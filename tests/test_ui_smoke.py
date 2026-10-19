"""Launch the complete Streamlit UI to catch runtime widget/state errors."""
import unittest
from pathlib import Path
from streamlit.testing.v1 import AppTest

APP=Path(__file__).resolve().parents[1]/"app.py"

class StudioSmokeTests(unittest.TestCase):
    def test_studio_startup(self):
        at=AppTest.from_file(str(APP), default_timeout=20).run()
        errors=[str(e.message) for e in at.exception]
        self.assertEqual(errors, [], "Streamlit startup errors: " + " | ".join(errors))
        self.assertGreaterEqual(len(at.tabs), 2)

    def test_quick_generate_export_button(self):
        at=AppTest.from_file(str(APP), default_timeout=30)
        at.session_state["raw_data"]=(
            "No.,Word,Type,Pronunciation,Meaning\n"
            "1,protect,v,/prəˈtekt/,bảo vệ\n"
            ",protection,n,/prəˈtekʃn/,sự bảo vệ\n"
        )
        at.run()
        self.assertEqual(len(at.exception),0)
        buttons=[button for button in at.button if button.key=="studio_quick_generate"]
        self.assertEqual(len(buttons),1)
        buttons[0].click().run()
        self.assertEqual(len(at.exception),0)
        self.assertTrue(at.session_state["_last_export_bytes"])

if __name__ == "__main__":
    unittest.main()
