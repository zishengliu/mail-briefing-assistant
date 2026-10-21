import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from email import policy
from email.message import EmailMessage
from pathlib import Path

import mail_briefing as app


class MailAssistantTests(unittest.TestCase):
    def setUp(self):
        self.account = {
            "id": "gmail",
            "provider_id": "gmail",
            "name": "Gmail",
            "host": "imap.gmail.com",
            "port": 993,
            "url": "https://mail.google.com/",
            "email": "test@gmail.com",
        }

    def test_plain_text_body_and_priority(self):
        message = EmailMessage()
        message["From"] = "Office <office@example.edu>"
        message["To"] = "test@example.com"
        message["Subject"] = "Deadline: please reply today"
        message["Date"] = datetime.now(timezone.utc).strftime("%a, %d %b %Y %H:%M:%S +0000")
        message.set_content("Please review the attached material.\nThe deadline is today.")
        mail = app.parse_message(message.as_bytes(policy=policy.SMTP), 11, self.account)
        self.assertIn("deadline", mail["preview"].lower())
        self.assertEqual(mail["category"], "需尽快处理")
        self.assertEqual(mail["priority"], "高")

    def test_html_fallback_extracts_visible_text(self):
        message = EmailMessage()
        message["From"] = "Notice <notice@example.com>"
        message["Subject"] = "A newsletter"
        message.set_content(
            "<html><head><style>.x { color: red; }</style></head>"
            "<body><script>secret_script_text</script><p>Hello <b>world</b></p>"
            "<div>Readable message body</div></body></html>",
            subtype="html",
        )
        text = app.extract_body(message, limit=1000)
        self.assertIn("Hello", text)
        self.assertIn("world", text)
        self.assertIn("Readable message body", text)
        self.assertNotIn("secret_script_text", text)
        self.assertNotIn("color: red", text)

    def test_old_categories_are_migrated(self):
        record = {
            "uid": 1, "account_id": "qq", "account_name": "QQ 邮箱",
            "category": "紧急/待处理", "subject": "Urgent", "sender": "a",
            "date": datetime.now(timezone.utc).isoformat(),
        }
        restored = app._restore_mail(record)
        self.assertEqual(restored["category"], "需尽快处理")
        self.assertEqual(restored["priority"], "高")

    def test_only_local_calendar_day_is_visible(self):
        today = datetime.now().astimezone().date()
        current = {"date": datetime.now().astimezone(), "subject": "today"}
        yesterday = {"date": datetime.now().astimezone() - timedelta(days=1), "subject": "old"}
        unknown = {"date": None, "subject": "unknown"}
        filtered = app.today_only([current, yesterday, unknown], today)
        self.assertEqual([item["subject"] for item in filtered], ["today"])

    def test_custom_imap_account_survives_settings_roundtrip(self):
        old_path = app.SETTINGS_PATH
        try:
            with tempfile.TemporaryDirectory() as directory:
                app.SETTINGS_PATH = Path(directory) / "settings.json"
                settings = app.default_settings()
                settings["accounts"].append({
                    "id": "custom-example123",
                    "provider_id": "custom",
                    "name": "Example Mail",
                    "host": "imap.example.com",
                    "port": 993,
                    "url": "https://mail.example.com/",
                    "email": "me@example.com",
                    "username": "imap-login",
                    "enabled": True,
                })
                app.save_settings(settings)
                loaded = app.load_settings()
                custom = next(a for a in loaded["accounts"] if a["id"] == "custom-example123")
                self.assertEqual(custom["host"], "imap.example.com")
                self.assertEqual(custom["port"], 993)
                self.assertEqual(custom["url"], "https://mail.example.com/")
                self.assertEqual(custom["username"], "imap-login")
        finally:
            app.SETTINGS_PATH = old_path


if __name__ == "__main__":
    unittest.main()
