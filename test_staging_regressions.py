import unittest
from pathlib import Path


class PartyManagementRegressionTests(unittest.TestCase):
    def test_party_list_uses_full_editor_and_guarded_delete(self):
        template = Path("module_manage.html").read_text(encoding="utf-8")
        self.assertIn("url_for('party_edit'", template)
        self.assertIn("url_for('party_delete'", template)

    def test_party_view_uses_full_editor(self):
        template = Path("module_view.html").read_text(encoding="utf-8")
        self.assertIn("url_for('party_edit'", template)


if __name__ == "__main__":
    unittest.main()
