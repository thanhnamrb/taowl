"""Title layout must fit reliably across fonts and Word/LibreOffice engines."""
import io
import unittest

from docx import Document
from docx.oxml.ns import qn
from title_fitter import fit_title, measured_width_pt

from test_original_template import ROOT, namespace


class FitterUnitTests(unittest.TestCase):
    CELL_WIDTH_PT = 5.138 * 72

    def test_people_and_places_is_one_line(self):
        result = fit_title("UNIT 7: PEOPLE & PLACES", cell_width_pt=self.CELL_WIDTH_PT)
        self.assertEqual(len(result.lines), 1, result)
        self.assertLessEqual(result.largest_line_width_pt, result.available_width_pt)
        self.assertGreaterEqual(result.font_size, 18)

    def test_many_realistic_titles(self):
        titles = [
            "UNIT 5: SPECIAL DAYS",
            "UNIT 7: PEOPLE & PLACES",
            "UNIT 10: ENVIRONMENTAL PROTECTION",
            "MODULE 12: FOOD, HEALTH & CULTURE",
            "UNIT 6: WHAT MAKES A GOOD LEADER?",
            "UNIT 14: GLOBAL COMMUNICATION AND TECHNOLOGY",
            "UNIT 8: SCIENTIFIC DISCOVERIES THROUGH HISTORY",
            "UNIT 100: THE LAST CHAPTER",
            "UNIT 5: THIS-IS-A-VERY-LONG-UNBREAKABLE-TOKEN-IN-THE-TITLE",
            "UNIT 11: ADVANCED PROBLEM SOLVING IN REAL-LIFE SITUATIONS",
            "UNIT 16: THE EFFECTS OF SOCIAL MEDIA ON HUMAN INTERACTION",
            "UNIT 3: EXTRAORDINARY BIOLOGICAL DIVERSITY ACROSS CONTINENTS",
            "UNIT 22: " + "CULTURE SCIENCE HISTORY TECHNOLOGY " * 4,
        ]
        for title in titles:
            with self.subTest(title=title):
                result = fit_title(title, cell_width_pt=self.CELL_WIDTH_PT)
                self.assertEqual("".join(result.text.split()), "".join(title.split()))
                self.assertTrue(all(
                    measured_width_pt(line, result.font_size) <= result.available_width_pt + .1
                    for line in result.lines
                ))
                self.assertTrue(len(result.lines) <= 10)

    def test_invalid_width_or_empty_title_is_rejected(self):
        with self.assertRaises(ValueError):
            fit_title(" ", cell_width_pt=self.CELL_WIDTH_PT)
        with self.assertRaises(ValueError):
            fit_title("UNIT 1: GOOD", cell_width_pt=40)

    def test_export_has_fixed_grid_and_explicit_breaks(self):
        for title in ["PEOPLE & PLACES", "TECHNOLOGY AND OUR FUTURE WORLD",
                      "VERY LONG EDUCATIONAL AND SCIENTIFIC DISCOVERIES ACROSS THE WORLD"]:
            with self.subTest(title=title):
                doc = namespace["parse_vocab_csv"](
                    "1,protect,v,/prəˈtekt/,bảo vệ",
                    unit="7.1",heading_unit="7",title=title,
                )
                result = Document(io.BytesIO(
                    namespace["render_vocab_docx"](
                        doc, template_path=ROOT/"templates"/"lx_vocab_template.docx"
                    ).getvalue()
                ))
                table = result.tables[0]
                self.assertFalse(table.autofit)
                self.assertEqual(table._tbl.tblPr.find(qn("w:tblLayout")).get(qn("w:type")), "fixed")
                title_cell = table.rows[1].cells[2]
                self.assertIsNotNone(title_cell._tc.tcPr.find(qn("w:noWrap")))
                self.assertEqual(
                    " ".join(title_cell.text.split()),
                    "UNIT 7: " + title,
                )


if __name__ == "__main__":
    unittest.main()
