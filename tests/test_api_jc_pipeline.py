"""Regression tests for the API_JC repair (titles, BUSCARV, dups, groups, labels)."""
from __future__ import annotations

import io
import json
import time
import unittest
from copy import deepcopy

from openpyxl import Workbook, load_workbook

import api_jc_core as core


VIDEO_TITLE = (
    "Video | Topos Azteca destacan la solidaridad de los caleños "
    "durante las labores de rescate"
)

BUSCARV_FIXTURE = [
    ("el pais", "Nacional"),
    ("caracol radio", "Bogotá"),
    ("rcn radio", "Bogotá"),
    ("semana", "Nacional"),
    ("el tiempo", "Nacional"),
]
BUSCARV_INTERNET = [
    ("elpais.com.co", "El País"),
    ("semana.com", "Semana"),
    ("eltiempo.com", "El Tiempo"),
]


def _row(**kw):
    base = {
        "ID Noticia": kw.pop("id", "1"),
        "Fecha": kw.pop("fecha", "2026-01-15"),
        "Hora": kw.pop("hora", "08:00"),
        "Medio": kw.pop("medio", "El País"),
        "Tipo de Medio": kw.pop("tipo", "Internet"),
        "Sección - Programa": "Nacional",
        "Región": "Nacional",
        "Título": kw.pop("titulo", VIDEO_TITLE),
        "_titulo_original": kw.pop("titulo_original", None),
        "Tono IA": "",
        "Tema": "",
        "Subtema": "",
        "Link Nota": kw.pop("link", {"value": "Link", "url": "https://elpais.com.co/a"}),
        "Resumen - Aclaracion": kw.pop("resumen", "Resumen de la nota."),
        "Link (Streaming - Imagen)": None,
        "Menciones - Empresa": kw.pop("mencion", "Topos Azteca"),
        "ID duplicada": "",
        "Cuerpo Completo": kw.pop("cuerpo", ""),
        "Tono": "Neutro",
        "is_duplicate": False,
    }
    if base["_titulo_original"] is None:
        base["_titulo_original"] = base["Título"]
    base.update(kw)
    return base


class TestTitleUnchanged(unittest.TestCase):
    def test_01_video_pipe_title_byte_for_byte(self):
        rows = [_row(id="100", titulo=VIDEO_TITLE)]
        out, _ = core.process_pipeline(rows, "Topos Azteca", ["Topos"])
        self.assertEqual(out[0]["Título"], VIDEO_TITLE)
        xlsx = core.generate_output_excel(out)
        wb = load_workbook(io.BytesIO(xlsx))
        headers = [c.value for c in wb.active[1]]
        col = headers.index("Título") + 1
        self.assertEqual(wb.active.cell(2, col).value, VIDEO_TITLE)
        self.assertNotIn("|", core.normalize_title_for_comparison(VIDEO_TITLE) or "|")
        # Internal normalize may drop punctuation, output must not.
        self.assertTrue(VIDEO_TITLE.startswith("Video |"))


class TestBuscarv(unittest.TestCase):
    def test_02_buscarv_fixture_identical_mapping(self):
        region_map = core.construir_mapa_buscarv(BUSCARV_FIXTURE)
        internet_map = core.construir_mapa_buscarv(BUSCARV_INTERNET)
        self.assertEqual(core.aplicar_buscarv("El Pais", region_map), "Nacional")
        self.assertEqual(core.aplicar_buscarv("CARACOL RADIO", region_map), "Bogotá")
        self.assertEqual(core.aplicar_buscarv("medio-fantasma", region_map), "N/A")
        regiones, medios = core.aplicar_buscarv_dossier(
            ["El Pais", "elpais.com.co", "Caracol Radio"],
            region_map, internet_map,
            tipos=["Internet", "Internet", "Radio"],
        )
        self.assertEqual(regiones, ["Nacional", "N/A", "Bogotá"])
        self.assertEqual(medios[0], "El Pais")  # no Internet alias for this key
        self.assertEqual(medios[1], "El País")  # BUSCARV Internet
        self.assertEqual(medios[2], "Caracol Radio")  # radio unchanged
        # Control/branch name remains BUSCARV
        self.assertTrue(hasattr(core, "aplicar_buscarv"))
        self.assertTrue(hasattr(core, "construir_mapa_buscarv"))


class TestDuplicatesVsGrupo(unittest.TestCase):
    def test_03_internet_same_title_different_urls_not_duplicate(self):
        rows = [
            _row(id="10", titulo=VIDEO_TITLE, medio="El País",
                 link={"value": "Link", "url": "https://elpais.com.co/solidaridad-cali"}),
            _row(id="11", titulo=VIDEO_TITLE, medio="Semana",
                 link={"value": "Link", "url": "https://semana.com/solidaridad-cali"}),
        ]
        out, _ = core.process_pipeline(rows, "Topos Azteca", ["Topos"])
        self.assertFalse(out[0].get("is_duplicate"))
        self.assertFalse(out[1].get("is_duplicate"))
        self.assertEqual(out[1].get("ID duplicada") or "", "")
        self.assertEqual(out[0]["Grupo noticia"], out[1]["Grupo noticia"])
        self.assertEqual(out[0]["Tono IA"], out[1]["Tono IA"])
        self.assertEqual(out[0]["Tema"], out[1]["Tema"])
        self.assertEqual(out[0]["Subtema"], out[1]["Subtema"])
        self.assertNotEqual(out[0]["Tono IA"], "Duplicada")

    def test_04_internet_same_embedded_url_is_duplicate(self):
        url = "https://WWW.ElPais.com.co/nota/abc/?utm_source=tw&utm_campaign=x#frag"
        url2 = "https://elpais.com.co/nota/abc/"
        rows = [
            _row(id="20", titulo="Primera publicación",
                 link={"value": "Link", "url": url}),
            _row(id="21", titulo="Republicación distinta",
                 link={"value": "Link", "url": url2}),
        ]
        out = core.detectar_duplicados(rows)
        self.assertFalse(out[0].get("is_duplicate"))
        self.assertTrue(out[1].get("is_duplicate"))
        self.assertEqual(str(out[1]["ID duplicada"]), "20")
        self.assertEqual(out[1]["Tono IA"], "Duplicada")
        xlsx = core.generate_output_excel(out)
        wb = load_workbook(io.BytesIO(xlsx))
        headers = [c.value for c in wb.active[1]]
        link_col = headers.index("Link Nota") + 1
        cell = wb.active.cell(2, link_col)
        self.assertTrue(cell.hyperlink)
        self.assertIn("elpais.com.co/nota/abc", cell.hyperlink.target.lower())

    def test_05_radio_tv_same_datetime_duplicate_different_not(self):
        title_a = "Alcalde presenta balance de seguridad en Cali"
        title_b = "Alcalde presenta balance de seguridad en Cali esta noche"
        rows = [
            _row(id="30", tipo="Radio", medio="Caracol Radio", fecha="2026-03-01",
                 hora="07:00", titulo=title_a, link=None),
            _row(id="31", tipo="Televisión", medio="Telepacífico", fecha="2026-03-01",
                 hora="07:00", titulo=title_b, link=None),
            _row(id="32", tipo="Radio", medio="RCN Radio", fecha="2026-03-01",
                 hora="09:30", titulo=title_a, link=None),
            _row(id="33", tipo="Radio", medio="Caracol Radio", fecha="2026-03-02",
                 hora="07:00", titulo=title_a, link=None),
        ]
        out = core.detectar_duplicados(rows)
        self.assertFalse(out[0]["is_duplicate"])
        self.assertTrue(out[1]["is_duplicate"], "same title+fecha+hora across media is duplicate")
        self.assertEqual(str(out[1]["ID duplicada"]), "30")
        self.assertFalse(out[2]["is_duplicate"], "different Hora is not duplicate")
        self.assertFalse(out[3]["is_duplicate"], "different Fecha is not duplicate")


class TestGroupingAndLabels(unittest.TestCase):
    def test_06_similar_resumen_same_grupo_and_labels(self):
        r1 = (
            "Las autoridades confirmaron el aumento de heridos tras el "
            "terremoto que afectó el suroccidente del país."
        )
        r2 = (
            "Las autoridades confirmaron el aumento de heridos tras el "
            "terremoto que sacudió el suroccidente colombiano."
        )
        rows = [
            _row(id="40", titulo="Balance oficial tras el sismo", resumen=r1,
                 link={"value": "Link", "url": "https://a.com/1"}),
            _row(id="41", titulo="Nuevas cifras del movimiento telúrico", resumen=r2,
                 link={"value": "Link", "url": "https://b.com/2"}),
        ]
        out, _ = core.process_pipeline(rows, "Cruz Roja", ["Cruz Roja Colombiana"])
        self.assertEqual(out[0]["Grupo noticia"], out[1]["Grupo noticia"])
        self.assertEqual(out[0]["Tono IA"], out[1]["Tono IA"])
        self.assertEqual(out[0]["Tema"], out[1]["Tema"])
        self.assertEqual(out[0]["Subtema"], out[1]["Subtema"])
        self.assertFalse(out[0]["is_duplicate"])
        self.assertFalse(out[1]["is_duplicate"])

    def test_07_bad_labels_rejected_grammatical_replacements(self):
        bad = [
            "Terremoto en colombia ascienden",
            "Sismo en cali cómo saber si",
        ]
        for et in bad:
            self.assertFalse(core.validar_subtema(et), et)
        ctx1 = "Tras el terremoto en Colombia ascienden las víctimas y los heridos."
        ctx2 = "Sismo en Cali: cómo saber si la vivienda quedó bien y qué recomendaciones seguir."
        s1 = core.fallback_subtema(ctx1, "Terremoto en colombia ascienden")
        s2 = core.fallback_subtema(ctx2, "Sismo en cali cómo saber si")
        self.assertTrue(core.validar_subtema(s1), s1)
        self.assertTrue(core.validar_subtema(s2), s2)
        self.assertNotIn("ascienden", s1.lower())
        self.assertNotIn("cómo saber", s2.lower())
        self.assertIn("víctimas", s1.lower())
        self.assertTrue("sismo" in s2.lower() or "cali" in s2.lower())

    def test_08_diporto_no_keyword_collage(self):
        collage = "Diporto resultados jornada fecha liga"
        self.assertFalse(core.validar_subtema(collage))
        sub = core.fallback_subtema(
            "Diporto entregó los resultados de la jornada y la fecha de liga.",
            collage,
        )
        self.assertTrue(core.validar_subtema(sub), sub)
        self.assertNotEqual(core._norm_text(sub), core._norm_text(collage))
        self.assertTrue(any(n in sub.lower() for n in ("de", "del", "en", "sobre")))


class TestOutputSchema(unittest.TestCase):
    def test_09_exact_output_columns_no_extras(self):
        rows = [_row(id="50", titulo=VIDEO_TITLE)]
        out, _ = core.process_pipeline(rows, "Topos Azteca")
        xlsx = core.generate_output_excel(out)
        wb = load_workbook(io.BytesIO(xlsx))
        headers = [c.value for c in wb.active[1]]
        self.assertEqual(headers, core.OUTPUT_COLUMNS)
        self.assertNotIn("Autor - Conductor", headers)
        self.assertNotIn("CPE", headers)
        self.assertNotIn("Nro. Pagina", headers)


class TestScaleAndBatching(unittest.TestCase):
    def test_10_400_rows_blocked_not_n2_bounded_chat(self):
        n = 400
        rows = []
        for i in range(n):
            rows.append(_row(
                id=str(1000 + i),
                titulo=f"Zalpha{i:04d} Quorum{i:04d} hallazgo local independiente",
                resumen=f"Resumen exclusivo ZX{i:04d} sobre un hecho puntual distinto.",
                fecha="2026-04-01",
                hora=f"{(i % 20) + 1:02d}:00",
                medio=f"Medio{i}",
                link={"value": "Link", "url": f"https://noticias.example/{i}/nota-{i}"},
                mencion="MarcaX",
            ))
        # Two extra rows that SHOULD group (similar resumen) and one URL dup.
        rows[10]["Resumen - Aclaracion"] = rows[11]["Resumen - Aclaracion"] = (
            "Las autoridades confirmaron el aumento de heridos tras el terremoto andino."
        )
        rows[10]["Título"] = "Balance oficial del terremoto andino"
        rows[11]["Título"] = "Cifras actualizadas del terremoto andino"
        rows[20]["Link Nota"] = deepcopy(rows[21]["Link Nota"])

        chat_calls = []

        def chat_fn(prompt):
            chat_calls.append(prompt)
            # Count items in this batch.
            try:
                start = prompt.index("ÍTEMS:")
                payload = json.loads(prompt[start + 6:])
            except Exception:
                payload = []
            items = []
            for it in payload:
                items.append({
                    "id": it["id"],
                    "tono": "Neutro",
                    "tema": "Cobertura informativa",
                    "subtema": "Cobertura de información local",
                })
            return {"items": items}

        embed_calls = []

        def embed_fn(textos):
            embed_calls.append(len(textos))
            rng = __import__("numpy").random.RandomState(0)
            return [rng.randn(16) for _ in textos]

        t0 = time.time()
        out, prog = core.process_pipeline(
            rows, "MarcaX", embed_fn=embed_fn, chat_fn=chat_fn,
        )
        cpu = time.time() - t0
        self.assertLess(cpu, 8.0, f"grouping/pipeline CPU too slow: {cpu:.2f}s")
        self.assertEqual(len(embed_calls), 1, "must be a single embedding pass")
        self.assertEqual(embed_calls[0], n)
        max_pairs = n * (n - 1) // 2
        self.assertLess(prog.comparisons.n, n * 25)
        self.assertLess(prog.comparisons.n, max_pairs // 4)
        self.assertLessEqual(len(chat_calls), (n + core.CHAT_BATCH_SIZE - 1) // core.CHAT_BATCH_SIZE + 2)
        self.assertGreaterEqual(len(chat_calls), 1)
        for prompt in chat_calls:
            # Each request is a batch, not one news item.
            self.assertIn("ÍTEMS:", prompt)
        dups = [r for r in out if r.get("is_duplicate")]
        self.assertTrue(any(str(r.get("ID duplicada")) == str(rows[20]["ID Noticia"]) or
                            str(r.get("ID duplicada")) == str(rows[21]["ID Noticia"])
                            for r in dups))
        self.assertEqual(out[10]["Grupo noticia"], out[11]["Grupo noticia"])
        stages = [e["stage"] for e in prog.events]
        for required in ("Limpieza", "Duplicados", "Contexto", "Embedding único",
                         "Agrupación", "Tono", "Tema/Subtema"):
            self.assertIn(required, stages)


class TestUrlAndHyperlinkHelpers(unittest.TestCase):
    def test_url_normalization_keeps_distinct_paths(self):
        a = core.normalize_url("https://WWW.Site.com/a/b/?utm_source=x#z")
        b = core.normalize_url("http://site.com/a/b")
        c = core.normalize_url("https://site.com/a/c/")
        self.assertEqual(a, b)
        self.assertNotEqual(a, c)

    def test_hyperlink_roundtrip_from_workbook(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "Link Nota"
        ws["B1"] = "Título"
        cell = ws["A2"]
        cell.value = "Link"
        cell.hyperlink = "https://ejemplo.com/nota-1"
        ws["B2"] = VIDEO_TITLE
        buf = io.BytesIO()
        wb.save(buf)
        parsed = core.read_titulo_and_links_from_xlsx(buf.getvalue())
        self.assertEqual(parsed[0]["Título"], VIDEO_TITLE)
        self.assertEqual(core.url_de_celda_link(parsed[0]["Link Nota"]),
                         "https://ejemplo.com/nota-1")


if __name__ == "__main__":
    unittest.main()
