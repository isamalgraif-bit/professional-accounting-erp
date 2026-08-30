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

<<<<<<< Updated upstream
=======
    def test_party_editor_loads_contact_values_by_mapping_key(self):
        template = Path("party_edit.html").read_text(encoding="utf-8")
        self.assertIn("party['phone']", template)
        self.assertIn("party['email']", template)

    def test_party_update_preserves_blank_contact_submissions(self):
        source = Path("app.py").read_text(encoding="utf-8")
        self.assertIn("phone=COALESCE(NULLIF(:phone,''),phone)", source)
        self.assertIn("email=COALESCE(NULLIF(:email,''),email)", source)


class StagingUiRegressionTests(unittest.TestCase):
    def test_dashboard_breadcrumb_uses_flask_dashboard_route(self):
        template = Path("base.html").read_text(encoding="utf-8")
        self.assertNotIn('href="/dashboard"', template)
        self.assertIn('url_for("dashboard")', template)

    def test_autosave_toast_is_only_defined_on_real_autosave_editors(self):
        base = Path("base.html").read_text(encoding="utf-8")
        self.assertIn("{% block auto_save_toast %}{% endblock %}", base)
        for name in ("journal_entries.html", "multi_journal.html"):
            self.assertIn('id="globalSaveToast"', Path(name).read_text(encoding="utf-8"))
        self.assertNotIn('id="globalSaveToast"', base)

    def test_data_import_uses_central_application_version(self):
        template = Path("data_import_center.html").read_text(encoding="utf-8")
        self.assertIn("{{app_version}}", template)
        self.assertNotIn("20.0.3", template)

    def test_multi_journal_has_one_class_attribute_and_group_totals(self):
        template = Path("multi_journal.html").read_text(encoding="utf-8")
        self.assertIn('class="table table-bordered table-sm align-middle multi-entry-grid"', template)
        self.assertIn('id="groupSummaries"', template)
        self.assertIn("مدين:", template)
        self.assertIn("دائن:", template)
        self.assertIn("الفرق:", template)
        self.assertIn("متوازن", template)

    def test_workspace_exposes_full_open_screen_list(self):
        template = Path("base.html").read_text(encoding="utf-8")
        self.assertIn("erp-workspace-list-menu", template)
        self.assertIn("كل الشاشات (", template)

    def test_executive_dashboard_isolates_optional_widget_queries(self):
        source = Path("app.py").read_text(encoding="utf-8")
        route = source[source.index('def executive_dashboard():'):source.index('@app.route("/reports")')]
        self.assertIn("bi_safe_row", route)
        self.assertIn("bi_safe_rows", route)
        self.assertNotIn("=row(", route)
        self.assertNotIn("=rows(", route)

>>>>>>> Stashed changes

if __name__ == "__main__":
    unittest.main()
