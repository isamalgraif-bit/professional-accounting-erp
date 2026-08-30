import unittest
from pathlib import Path
from unittest.mock import patch

import app as erp


class PartyManagementRegressionTests(unittest.TestCase):
    def test_party_list_uses_full_editor_and_guarded_delete(self):
        template = Path("module_manage.html").read_text(encoding="utf-8")
        self.assertIn("url_for('party_edit'", template)
        self.assertIn("url_for('party_delete'", template)

    def test_party_view_uses_full_editor(self):
        template = Path("module_view.html").read_text(encoding="utf-8")
        self.assertIn("url_for('party_edit'", template)

    def test_party_editor_loads_contact_values_by_mapping_key(self):
        template = Path("party_edit.html").read_text(encoding="utf-8")
        self.assertIn("party['phone']", template)
        self.assertIn("party['email']", template)
        self.assertIn("party['party_account_id']|string==a['id']|string", template)

    def test_party_update_preserves_blank_contact_submissions(self):
        source = Path("app.py").read_text(encoding="utf-8")
        self.assertIn("phone=COALESCE(NULLIF(:phone,''),phone)", source)
        self.assertIn("email=COALESCE(NULLIF(:email,''),email)", source)


class PartyEditRouteTests(unittest.TestCase):
    def setUp(self):
        erp.app.config.update(TESTING=True, SECRET_KEY="test")
        erp.app._db_initialized = True
        self.party = {
            "id": 8, "name": "عميل اختبار", "name_en": "QA Customer",
            "vat_number": "", "phone": "0500000001",
            "email": "qa-customer-updated@example.invalid",
            "party_account_id": 113,
        }
        self.account = {"id": 113, "account_code": "1130", "account_name_ar": "عميل اختبار"}

    def _row(self, query, params=None):
        if "FROM customers" in query or "FROM suppliers" in query:
            if "id<>:id" in query:
                return None
            return dict(self.party)
        if "FROM chart_of_accounts" in query:
            return dict(self.account)
        if "FROM settings" in query:
            return {}
        return None

    def _get(self, party_type):
        with erp.app.test_client() as client, \
             patch.object(erp, "row", side_effect=self._row), \
             patch.object(erp, "rows", return_value=[dict(self.account)]), \
             patch.object(erp, "has_permission", return_value=True):
            with client.session_transaction() as session:
                session["user_id"] = 1
            return client.get(f"/parties/{party_type}/8/edit")

    def test_customer_get_renders_current_contact_and_account(self):
        response = self._get("customer")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('name="phone" value="0500000001"', html)
        self.assertIn('name="email" value="qa-customer-updated@example.invalid"', html)
        self.assertIn('value="113" data-account-name="عميل اختبار" selected', html)

    def test_supplier_get_uses_the_same_mapping_safely(self):
        response = self._get("supplier")
        self.assertEqual(response.status_code, 200)
        self.assertIn('name="phone" value="0500000001"', response.get_data(as_text=True))

    def _post(self, form):
        executed = []
        with erp.app.test_request_context("/parties/customer/8/edit", method="POST", data=form), \
             patch.object(erp, "row", side_effect=self._row), \
             patch.object(erp, "execute", side_effect=lambda query, params=None: executed.append((query, params))), \
             patch.object(erp, "audit"):
            response = erp.party_edit.__wrapped__("customer", 8)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(len(executed), 1)
        return executed[0][1]

    def test_post_preserves_contacts_and_account_when_not_changed(self):
        params = self._post({"name": "عميل اختبار", "name_en": "QA Customer"})
        self.assertEqual(params["phone"], self.party["phone"])
        self.assertEqual(params["email"], self.party["email"])
        self.assertEqual(params["party_account_id"], self.party["party_account_id"])

    def test_post_saves_explicit_contact_changes_without_changing_account(self):
        params = self._post({
            "name": "عميل اختبار", "name_en": "QA Customer",
            "phone": "0555555555", "email": "changed@example.invalid",
        })
        self.assertEqual(params["phone"], "0555555555")
        self.assertEqual(params["email"], "changed@example.invalid")
        self.assertEqual(params["party_account_id"], self.party["party_account_id"])


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

if __name__ == "__main__":
    unittest.main()
