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

if __name__ == "__main__":
    unittest.main()
