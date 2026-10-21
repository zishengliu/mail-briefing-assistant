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

    def test_connection_recovery_advice_explains_authentication_fix(self):
        account = {
            "id": "tsinghua", "provider_id": "tsinghua",
            "name": "清华大学邮箱", "email": "test@tsinghua.edu.cn",
        }
        guidance = app.connection_recovery_advice(account, "LOGIN failed: AUTHENTICATIONFAILED")
        self.assertIn("客户端专用密码/授权码", guidance)
        self.assertIn("开启 IMAP", guidance)
        self.assertIn("清华大学邮箱：连接失败", guidance)

    def test_mailbox_status_labels_show_counts_and_connection_failures(self):
        accounts = [
            {"id": "gmail", "name": "Gmail", "email": "test@gmail.com", "enabled": True},
            {"id": "tsinghua", "name": "清华大学邮箱", "email": "test@tsinghua.edu.cn", "enabled": True},
            {"id": "qq", "name": "QQ 邮箱", "email": "test@qq.com", "enabled": True},
            {"id": "off", "name": "已停用邮箱", "email": "off@example.com", "enabled": False},
        ]
        mails = [
            {"account_id": "gmail"}, {"account_id": "gmail"}, {"account_id": "qq"},
        ]
        errors = ["清华大学邮箱：连接失败（timeout）\\n恢复建议：重新测试连接。"]
        labels = app.mailbox_status_labels(mails, accounts, errors)
        self.assertEqual(labels, ["Gmail 2 封", "清华大学邮箱 连接异常", "QQ 邮箱 1 封"])
        self.assertFalse(any("已停用邮箱" in label for label in labels))

    def test_daily_schedule_defaults_to_eight_and_seventeen(self):
        settings = app.default_settings()
        self.assertEqual(app.configured_daily_times(settings), ["08:00", "17:00"])

    def test_legacy_single_time_keeps_custom_morning_and_adds_evening(self):
        old_path = app.SETTINGS_PATH
        try:
            with tempfile.TemporaryDirectory() as directory:
                app.SETTINGS_PATH = Path(directory) / "settings.json"
                app.SETTINGS_PATH.write_text('{"daily_time": "07:45", "accounts": []}', encoding="utf-8")
                settings = app.load_settings()
                self.assertEqual(app.configured_daily_times(settings), ["07:45", "17:00"])
        finally:
            app.SETTINGS_PATH = old_path

    def test_stored_mail_keeps_sender_address_for_original_link(self):
        stored = app._stored_mail({
            "uid": 9, "account_id": "gmail", "subject": "Meeting",
            "sender": "Office", "sender_address": "office@example.edu",
        })
        self.assertEqual(stored["sender_address"], "office@example.edu")

    def test_gmail_web_fallback_searches_by_subject(self):
        url = app.resolve_webmail_url(
            self.account,
            {"subject": "Meeting tomorrow", "sender_address": "notice@example.edu"},
        )
        self.assertIn("mail.google.com", url)
        self.assertIn("#search/", url)
        self.assertIn("Meeting", url)


    def test_outlook_web_fallback_searches_by_subject(self):
        account = {
            **self.account,
            "id": "outlook",
            "provider_id": "outlook",
            "url": "https://outlook.live.com/mail/",
        }
        url = app.resolve_webmail_url(account, {"subject": "Meeting tomorrow"})
        self.assertIn("outlook.live.com/mail/0/search/", url)
        self.assertIn("Meeting", url)

    def test_custom_mailbox_fallback_uses_configured_webmail_url(self):
        account = {
            **self.account,
            "id": "custom-1",
            "provider_id": "custom",
            "url": "https://mail.example.com/inbox",
        }
        self.assertEqual(
            app.resolve_webmail_url(account, {"subject": "Important notice"}),
            "https://mail.example.com/inbox",
        )

    def test_jump_to_original_opens_the_selected_message_search(self):
        from unittest.mock import patch

        message = {"subject": "Meeting tomorrow", "sender_address": "notice@example.edu"}
        with patch("mail_briefing.get_account_for_open", return_value=self.account):
            with patch("mail_briefing.webbrowser.open") as open_browser:
                app.open_mail_in_web("gmail", message)

        open_browser.assert_called_once()
        self.assertIn("#search/", open_browser.call_args.args[0])
        self.assertIn("Meeting", open_browser.call_args.args[0])

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
