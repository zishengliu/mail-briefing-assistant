#!/usr/bin/env python3
"""Mail Briefing Assistant: local IMAP inbox digest and Windows notifications."""
from __future__ import annotations

import argparse
import email
import imaplib
import json
import logging
import os
import re
import sys
import threading
import time
import webbrowser
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import keyring
except ImportError:
    keyring = None

APP_TITLE = "邮件晨报助手"
SERVICE_NAME = "MailBriefingAssistant"
PROVIDERS = [
    {"id": "gmail", "name": "Gmail", "host": "imap.gmail.com", "port": 993,
     "url": "https://mail.google.com/mail/u/0/#inbox", "help": "Gmail 建议使用 Google 官方授权；若使用 IMAP 应用密码，通常需要开启两步验证。"},
    {"id": "qq", "name": "QQ 邮箱", "host": "imap.qq.com", "port": 993,
     "url": "https://mail.qq.com/", "help": "请在邮箱设置中开启 IMAP/SMTP 服务，并填写生成的授权码（不是网页登录密码）。"},
    {"id": "tsinghua", "name": "清华大学邮箱（学生）", "host": "mails.tsinghua.edu.cn", "port": 993,
     "url": "https://mails.tsinghua.edu.cn/", "help": "学生邮箱 IMAP：mails.tsinghua.edu.cn:993；请填写网页版邮箱设置中的客户端专用密码。"},
]
CATEGORY_ORDER = ["紧急/待处理", "课程与学术", "求职与实习", "账单与账户", "订阅与通知", "个人与其他"]
RULES = {
    "紧急/待处理": ["紧急", "尽快", "请处理", "待处理", "待确认", "截止", "逾期", "deadline", "urgent", "action required", "asap", "due today", "approval required", "verify your account"],
    "课程与学术": ["清华", "北京大学", "课程", "教务", "研究生", "学术", "论文", "讲座", "seminar", "conference", "journal", "submission", "course", "university", "academic", "scholarship", "advisor", "导师", "答辩", "开题"],
    "求职与实习": ["招聘", "实习", "面试", "录用", "offer", "career", "recruit", "recruitment", "interview", "internship", "job opportunity", "快手", "美团", "简历", "岗位", "hr "],
    "账单与账户": ["账单", "发票", "付款", "支付", "退款", "订单", "invoice", "payment", "receipt", "billing", "subscription renewal", "bank statement", "密码重置", "password reset", "登录提醒", "security alert"],
    "订阅与通知": ["unsubscribe", "newsletter", "邮件订阅", "每周摘要", "weekly digest", "no-reply", "noreply", "notification", "通知摘要", "活动邀请", "产品更新", "product update", "digest"],
}


def user_data_dir() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home() / ".config")
    path = Path(base) / "MailBriefingAssistant"
    path.mkdir(parents=True, exist_ok=True)
    return path


DATA_DIR = user_data_dir()
SETTINGS_PATH = DATA_DIR / "settings.json"
STATE_PATH = DATA_DIR / "state.json"
DIGEST_PATH = DATA_DIR / "latest_digest.json"
LOG_PATH = DATA_DIR / "mail-briefing.log"

logging.basicConfig(filename=LOG_PATH, level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")


def default_settings() -> dict:
    return {
        "daily_time": "08:00",
        "accounts": [{"id": p["id"], "name": p["name"], "host": p["host"], "port": p["port"],
                      "url": p["url"], "email": "", "enabled": True} for p in PROVIDERS],
    }


def load_json(path: Path, fallback):
    try:
        with path.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, json.JSONDecodeError):
        return fallback


def load_settings() -> dict:
    result = default_settings()
    saved = load_json(SETTINGS_PATH, {})
    if isinstance(saved, dict):
        result["daily_time"] = saved.get("daily_time", result["daily_time"])
        saved_accounts = {a.get("id"): a for a in saved.get("accounts", []) if isinstance(a, dict)}
        for account in result["accounts"]:
            old = saved_accounts.get(account["id"], {})
            account["email"] = old.get("email", "")
            account["enabled"] = bool(old.get("enabled", True))
    return result


def save_settings(settings: dict) -> None:
    SETTINGS_PATH.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def load_state() -> dict:
    state = load_json(STATE_PATH, {})
    if not isinstance(state, dict):
        state = {}
    state.setdefault("last_uid", {})
    state.setdefault("last_digest_date", "")
    state.setdefault("pending_digest", [])
    return state


def save_state(state: dict) -> None:
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _stored_mail(mail: dict) -> dict:
    """Keep only the minimum metadata needed to show and notify a daily digest."""
    date_value = mail.get("date")
    if isinstance(date_value, datetime):
        date_value = date_value.isoformat()
    return {
        "uid": mail.get("uid"),
        "account_id": mail.get("account_id"),
        "account_name": mail.get("account_name", ""),
        "subject": mail.get("subject", "（无主题）"),
        "sender": mail.get("sender", "未知发件人"),
        "category": mail.get("category", "个人与其他"),
        "date": date_value,
    }


def _restore_mail(record: dict) -> dict:
    mail = dict(record)
    date_value = mail.get("date")
    if isinstance(date_value, str):
        try:
            mail["date"] = datetime.fromisoformat(date_value)
        except ValueError:
            mail["date"] = None
    return mail


def load_digest() -> dict:
    data = load_json(DIGEST_PATH, {})
    return data if isinstance(data, dict) else {}


def save_digest(mails: list[dict], errors: list[str]) -> None:
    payload = {
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "mails": [_stored_mail(mail) for mail in mails],
        "errors": errors,
    }
    DIGEST_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def get_secret(account_id: str) -> str | None:
    if keyring is None:
        return None
    try:
        return keyring.get_password(SERVICE_NAME, account_id)
    except Exception:
        logging.exception("Credential store lookup failed for %s", account_id)
        return None


def set_secret(account_id: str, secret: str) -> None:
    if keyring is None:
        raise RuntimeError("未安装 keyring，无法安全保存密码。请运行安装脚本或安装 requirements.txt。")
    keyring.set_password(SERVICE_NAME, account_id, secret)


def decode_header(value) -> str:
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def extract_body(message) -> str:
    parts = []
    try:
        if message.is_multipart():
            for part in message.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get_content_disposition() == "attachment":
                    continue
                if part.get_content_type() not in ("text/plain", "text/html"):
                    continue
                try:
                    value = part.get_content()
                except Exception:
                    payload = part.get_payload(decode=True) or b""
                    value = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                if part.get_content_type() == "text/html":
                    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", str(value))
                    value = re.sub(r"(?s)<[^>]+>", " ", str(value))
                parts.append(str(value))
                if sum(len(p) for p in parts) > 1600:
                    break
        else:
            try:
                value = message.get_content()
            except Exception:
                payload = message.get_payload(decode=True) or b""
                value = payload.decode(message.get_content_charset() or "utf-8", errors="replace")
            if message.get_content_type() == "text/html":
                value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", str(value))
                value = re.sub(r"(?s)<[^>]+>", " ", str(value))
            parts.append(str(value))
    except Exception:
        logging.debug("Could not extract message body", exc_info=True)
    text = re.sub(r"\s+", " ", " ".join(parts)).strip()
    return text[:450]


def parse_message(raw: bytes, uid: int, account: dict) -> dict:
    msg = BytesParser(policy=policy.default).parsebytes(raw)
    subject = decode_header(msg.get("subject")) or "（无主题）"
    from_header = decode_header(msg.get("from")) or "未知发件人"
    display, address = parseaddr(from_header)
    sender = display or address or from_header
    date_text = msg.get("date")
    received = None
    if date_text:
        try:
            received = parsedate_to_datetime(str(date_text))
            if received.tzinfo is None:
                received = received.replace(tzinfo=timezone.utc)
            received = received.astimezone(timezone.utc)
        except Exception:
            received = None
    return {"uid": uid, "account_id": account["id"], "account_name": account["name"],
            "account_email": account.get("email", ""), "subject": subject, "sender": sender,
            "sender_address": address or from_header, "date": received,
            "preview": extract_body(msg)}


def classify(mail: dict) -> str:
    haystack = (mail.get("subject", "") + " " + mail.get("sender", "") + " " +
                mail.get("sender_address", "") + " " + mail.get("preview", "")).lower()
    for category in CATEGORY_ORDER[:-1]:
        if any(term.lower() in haystack for term in RULES[category]):
            return category
    if any(x in haystack for x in ("edu.cn", "tsinghua.edu", "pku.edu", "course", "论文", "学术")):
        return "课程与学术"
    return "个人与其他"


def imap_connect(account: dict):
    if not account.get("email"):
        raise RuntimeError("尚未填写邮箱地址")
    secret = get_secret(account["id"])
    if not secret:
        raise RuntimeError("尚未保存客户端密码/授权码，请先在设置页填写并保存")
    conn = imaplib.IMAP4_SSL(account["host"], int(account.get("port", 993)), timeout=25)
    try:
        conn.login(account["email"], secret)
        typ, _ = conn.select("INBOX", readonly=True)
        if typ != "OK":
            raise RuntimeError("无法打开收件箱")
        return conn
    except Exception:
        try:
            conn.logout()
        except Exception:
            pass
        raise


def fetch_raw(conn, uid: int) -> bytes:
    status, data = conn.uid("fetch", str(uid), "(BODY.PEEK[]<0.4096>)")
    if status != "OK" or not data:
        raise RuntimeError(f"无法读取邮件 UID {uid}")
    chunks = [item[1] for item in data if isinstance(item, tuple) and isinstance(item[1], bytes)]
    if not chunks:
        raise RuntimeError(f"邮件 UID {uid} 未返回内容")
    return b"\r\n".join(chunks)


def check_account(account: dict, state: dict, test_only: bool = False) -> tuple[list[dict], int | None]:
    conn = imap_connect(account)
    try:
        status, data = conn.uid("search", None, "ALL")
        if status != "OK":
            raise RuntimeError("无法检索收件箱")
        all_uids = sorted(int(item) for item in (data[0] or b"").split() if item.isdigit())
        if test_only:
            return [], (all_uids[-1] if all_uids else 0)
        last_uid = state["last_uid"].get(account["id"])
        if last_uid is None:
            candidates = all_uids[-100:]
            initial_scan = True
        else:
            candidates = [uid for uid in all_uids if uid > int(last_uid)][:100]
            initial_scan = False
        results = []
        cutoff = datetime.now(timezone.utc) - timedelta(hours=24)
        for uid in candidates:
            raw = fetch_raw(conn, uid)
            item = parse_message(raw, uid, account)
            if initial_scan and item["date"] and item["date"] < cutoff:
                continue
            item["category"] = classify(item)
            results.append(item)
        if candidates:
            new_last = candidates[-1]
        else:
            new_last = all_uids[-1] if all_uids else (int(last_uid) if last_uid is not None else 0)
        return results, new_last
    finally:
        try:
            conn.logout()
        except Exception:
            pass


def send_notification(mails: list[dict], errors: list[str], digest_title: str = "邮件晨报") -> None:
    if not mails and not errors:
        logging.info("Daily check completed: no new messages")
        return
    counts = {category: sum(1 for mail in mails if mail["category"] == category) for category in CATEGORY_ORDER}
    count_line = " · ".join(f"{name} {count}" for name, count in counts.items() if count)
    details = [f"{mail['category']}｜{mail['sender']}: {mail['subject']}" for mail in mails[:3]]
    summary_parts = [part for part in (count_line, "\\n".join(details)) if part]
    if len(mails) > 3:
        summary_parts.append(f"另有 {len(mails) - 3} 封，请打开助手查看")
    if errors:
        summary_parts.append(f"{len(errors)} 个邮箱检查失败，请打开助手查看")
    summary = "\\n".join(summary_parts) or "没有发现新邮件"
    title = f"{digest_title}：发现 {len(mails)} 封新邮件" if mails else "邮件晨报：部分邮箱检查失败"
    try:
        from winotify import Notification
        Notification(app_id=APP_TITLE, title=title, msg=summary[:230], duration="short").show()
    except Exception as exc:
        logging.warning("Toast notification unavailable: %s", type(exc).__name__)
    logging.info("Check completed: %d new messages; %d account errors", len(mails), len(errors))


def do_check(settings: dict, state: dict, notify: bool = True, daily: bool = False) -> tuple[list[dict], list[str]]:
    all_mails, errors = [], []
    for account in settings["accounts"]:
        if not account.get("enabled") or not account.get("email"):
            continue
        try:
            mails, new_last_uid = check_account(account, state)
            state["last_uid"][account["id"]] = new_last_uid
            all_mails.extend(mails)
            logging.info("Account %s checked: %d new messages", account["id"], len(mails))
        except Exception as exc:
            message = f"{account['name']}：{str(exc)}"
            errors.append(message)
            logging.warning("Account %s failed: %s", account["id"], type(exc).__name__)

    # Manual checks queue the metadata for the next scheduled digest instead of consuming it.
    pending = list(state.get("pending_digest", []))
    seen = {(item.get("account_id"), item.get("uid")) for item in pending}
    for mail in all_mails:
        stored = _stored_mail(mail)
        key = (stored.get("account_id"), stored.get("uid"))
        if key not in seen:
            pending.append(stored)
            seen.add(key)
    state["pending_digest"] = pending
    save_state(state)

    if daily:
        digest_mails = [_restore_mail(item) for item in pending]
        save_digest(digest_mails, errors)
        if notify:
            send_notification(digest_mails, errors)
        state["pending_digest"] = []
        save_state(state)
    elif notify:
        send_notification(all_mails, errors)

    all_mails.sort(key=lambda m: m["date"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return all_mails, errors

def test_account(account: dict) -> str:
    conn = imap_connect(account)
    try:
        status, data = conn.uid("search", None, "ALL")
        if status != "OK":
            raise RuntimeError("已连接，但无法读取收件箱")
        count = len((data[0] or b"").split())
        return f"{account['name']}连接成功，收件箱约有 {count} 封邮件。"
    finally:
        try:
            conn.logout()
        except Exception:
            pass


def background_loop() -> None:
    """Wait for the configured local morning time, then produce one digest per day."""
    logging.info("Background scheduler started")
    last_minute = ""
    while True:
        try:
            now = datetime.now()
            settings = load_settings()
            state = load_state()
            minute_key = now.strftime("%Y-%m-%d %H:%M")
            due = datetime.strptime(settings.get("daily_time", "08:00"), "%H:%M").time()
            configured = any(a.get("enabled") and a.get("email") for a in settings["accounts"])
            if configured and now.time() >= due and state.get("last_digest_date") != now.date().isoformat() and minute_key != last_minute:
                do_check(settings, state, notify=True, daily=True)
                state = load_state()
                state["last_digest_date"] = now.date().isoformat()
                save_state(state)
                last_minute = minute_key
        except Exception as exc:
            logging.warning("Scheduler loop error: %s", type(exc).__name__)
        time.sleep(30)

class MailBriefingApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x690")
        self.root.minsize(760, 580)
        self.settings = load_settings()
        self.account_vars = {}
        self.password_vars = {}
        self.email_vars = {}
        self.enabled_vars = {}
        self.build_ui()
        self.render_saved_digest()

    def build_ui(self):
        wrapper = ttk.Frame(self.root, padding=14)
        wrapper.pack(fill="both", expand=True)
        ttk.Label(wrapper, text="邮件晨报助手", font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(wrapper, text="汇总 Gmail、QQ 邮箱与清华邮箱的新邮件；本机分类，不上传邮件内容。", foreground="#555").pack(anchor="w", pady=(2, 12))
        self.notebook = ttk.Notebook(wrapper)
        self.notebook.pack(fill="both", expand=True)
        self.settings_tab = ttk.Frame(self.notebook, padding=12)
        self.inbox_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.settings_tab, text="邮箱与定时设置")
        self.notebook.add(self.inbox_tab, text="邮件晨报")
        self.build_settings_tab()
        self.build_inbox_tab()
        bottom = ttk.Frame(wrapper)
        bottom.pack(fill="x", pady=(10, 0))
        self.status = tk.StringVar(value="请先配置邮箱并保存。")
        ttk.Label(bottom, textvariable=self.status).pack(side="left", fill="x", expand=True)
        ttk.Button(bottom, text="打开邮箱", command=self.open_webmail).pack(side="right", padx=(6, 0))
        ttk.Button(bottom, text="立即检查新邮件", command=self.check_now).pack(side="right")

    def build_settings_tab(self):
        ttk.Label(self.settings_tab, text="账号设置", font=("Microsoft YaHei UI", 11, "bold")).pack(anchor="w")
        for account in self.settings["accounts"]:
            provider = next(p for p in PROVIDERS if p["id"] == account["id"])
            card = ttk.LabelFrame(self.settings_tab, text=f"{account['name']} · {account['host']}:{account['port']}", padding=8)
            card.pack(fill="x", pady=5)
            enabled = tk.BooleanVar(value=account.get("enabled", True))
            email_var = tk.StringVar(value=account.get("email", ""))
            password_var = tk.StringVar()
            self.enabled_vars[account["id"]] = enabled
            self.email_vars[account["id"]] = email_var
            self.password_vars[account["id"]] = password_var
            ttk.Checkbutton(card, text="启用", variable=enabled).grid(row=0, column=0, sticky="w", padx=(0, 10))
            ttk.Label(card, text="邮箱地址").grid(row=0, column=1, sticky="e")
            ttk.Entry(card, textvariable=email_var, width=34).grid(row=0, column=2, sticky="ew", padx=5)
            ttk.Label(card, text="客户端密码 / 授权码").grid(row=1, column=1, sticky="e", pady=(5, 0))
            ttk.Entry(card, textvariable=password_var, width=34, show="•").grid(row=1, column=2, sticky="ew", padx=5, pady=(5, 0))
            ttk.Label(card, text=provider["help"], wraplength=690, foreground="#555").grid(row=2, column=1, columnspan=2, sticky="w", padx=5, pady=(5, 0))
            card.columnconfigure(2, weight=1)
        timing = ttk.LabelFrame(self.settings_tab, text="每日晨报", padding=8)
        timing.pack(fill="x", pady=(7, 4))
        ttk.Label(timing, text="通知时间（24 小时制）").grid(row=0, column=0, sticky="w")
        self.time_var = tk.StringVar(value=self.settings.get("daily_time", "08:00"))
        ttk.Entry(timing, textvariable=self.time_var, width=10).grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(timing, text="例如 08:00。安装启动任务后，电脑登录且助手运行时会按此时间检查。", foreground="#555").grid(row=0, column=2, sticky="w")
        row = ttk.Frame(self.settings_tab)
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(row, text="保存设置", command=self.save).pack(side="left")
        ttk.Button(row, text="测试已启用邮箱", command=self.test_all).pack(side="left", padx=8)
        ttk.Button(row, text="查看官方邮箱设置说明", command=lambda: webbrowser.open("https://mail.tsinghua.edu.cn/coremail/help/clientoption.jsp?locale=zh_CN")).pack(side="right")
        ttk.Label(self.settings_tab, text="密码仅保存在本机系统凭据库；设置文件只保存邮箱地址和同步状态。", foreground="#166534").pack(anchor="w", pady=(10, 0))

    def build_inbox_tab(self):
        toolbar = ttk.Frame(self.inbox_tab)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(toolbar, text="按分类整理新邮件", font=("Microsoft YaHei UI", 11, "bold")).pack(side="left")
        ttk.Button(toolbar, text="刷新列表", command=self.render_welcome).pack(side="right")
        columns = ("category", "account", "sender", "subject", "date")
        self.tree = ttk.Treeview(self.inbox_tab, columns=columns, show="headings", height=13)
        headings = {"category": "分类", "account": "邮箱", "sender": "发件人", "subject": "邮件主题", "date": "时间"}
        widths = {"category": 105, "account": 110, "sender": 140, "subject": 350, "date": 125}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], minwidth=75, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda event: self.open_webmail())
        self.summary = tk.StringVar(value="点击“立即检查新邮件”获取分类摘要。")
        ttk.Label(self.inbox_tab, textvariable=self.summary, wraplength=820, justify="left").pack(fill="x", pady=(8, 0))
        self.recent_mails = []

    def render_saved_digest(self):
        state = load_state()
        pending = [_restore_mail(item) for item in state.get("pending_digest", [])]
        if pending:
            self.show_results(pending, [], title="当前待晨报邮件", warn_errors=False)
            self.summary.set("当前待晨报邮件（这些邮件会在设定时间汇总通知）：\n" + self.summary.get())
            return
        digest = load_digest()
        if digest:
            mails = [_restore_mail(item) for item in digest.get("mails", [])]
            errors = digest.get("errors", [])
            self.show_results(mails, errors, title="上次晨报", warn_errors=False)
            generated_at = digest.get("generated_at", "时间未知")
            self.summary.set(f"上次晨报时间：{generated_at}\n" + self.summary.get())
            return
        self.render_welcome()

    def render_welcome(self):
        if hasattr(self, "tree"):
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.summary.set("最近一次检查结果会显示在这里。双击邮件可打开对应邮箱。")

    def save(self, quiet: bool = False):
        time_text = self.time_var.get().strip()
        try:
            datetime.strptime(time_text, "%H:%M")
        except ValueError:
            messagebox.showerror("时间格式错误", "请使用 24 小时制，例如 08:00。")
            return False
        settings = load_settings()
        settings["daily_time"] = time_text
        for account in settings["accounts"]:
            aid = account["id"]
            email_addr = self.email_vars[aid].get().strip()
            secret = self.password_vars[aid].get()
            account["email"] = email_addr
            account["enabled"] = self.enabled_vars[aid].get()
            if email_addr and "@" not in email_addr:
                messagebox.showerror("邮箱地址错误", f"{account['name']}：请输入完整的邮箱地址。")
                return False
            if secret:
                try:
                    set_secret(aid, secret)
                    self.password_vars[aid].set("")
                except Exception as exc:
                    messagebox.showerror("保存凭据失败", str(exc))
                    return False
        save_settings(settings)
        self.settings = settings
        self.status.set("设置已保存。首次检查会汇总近 24 小时的邮件，之后只处理新到邮件。")
        if not quiet:
            messagebox.showinfo("保存成功", "设置已保存。邮箱密码/授权码由本机系统凭据库保管。")
        return True

    def test_all(self):
        if not self.save(quiet=True):
            return
        accounts = [a for a in self.settings["accounts"] if a.get("enabled") and a.get("email")]
        if not accounts:
            messagebox.showwarning("尚未配置邮箱", "请至少启用一个邮箱并填写邮箱地址。")
            return
        self.status.set("正在测试邮箱连接……")
        def worker():
            results = []
            for account in accounts:
                try:
                    results.append("✓ " + test_account(account))
                except Exception as exc:
                    results.append("✗ " + account["name"] + "：" + str(exc))
            self.root.after(0, lambda: self.show_test_results(results))
        threading.Thread(target=worker, daemon=True).start()

    def show_test_results(self, results):
        self.status.set("邮箱连接测试完成。")
        messagebox.showinfo("连接测试结果", "\n\n".join(results))

    def check_now(self):
        if not self.save(quiet=True):
            return
        self.status.set("正在检查新邮件；邮件较多时可能需要一些时间……")
        def worker():
            state = load_state()
            _, errors = do_check(self.settings, state, notify=False, daily=False)
            updated_state = load_state()
            pending = [_restore_mail(item) for item in updated_state.get("pending_digest", [])]
            self.root.after(0, lambda: self.show_results(pending, errors, title="当前待晨报邮件"))
        threading.Thread(target=worker, daemon=True).start()

    def show_results(self, mails, errors, title="本次新邮件", warn_errors=True):
        self.recent_mails = mails
        for item in self.tree.get_children():
            self.tree.delete(item)
        counts = {c: 0 for c in CATEGORY_ORDER}
        for mail in mails:
            counts[mail["category"]] = counts.get(mail["category"], 0) + 1
            date_text = mail["date"].astimezone().strftime("%m-%d %H:%M") if mail["date"] else "未知"
            self.tree.insert("", "end", values=(mail["category"], mail["account_name"], mail["sender"], mail["subject"], date_text))
        self.summary.set(title + "：" + ("；".join(f"{k} {v} 封" for k, v in counts.items() if v) or "0 封") +
                         ("\n检查失败：" + "；".join(errors) if errors else "") +
                         ("\n分类采用本地关键词规则，不会调用外部 AI 服务。"))
        self.status.set(f"检查完成：{len(mails)} 封新邮件，{len(errors)} 个邮箱出错。")
        if errors and warn_errors:
            messagebox.showwarning("部分邮箱检查失败", "\n".join(errors))

    def open_webmail(self):
        selected = self.tree.selection()
        account_id = None
        if selected:
            values = self.tree.item(selected[0], "values")
            account_name = values[1] if len(values) > 1 else ""
            for a in self.settings["accounts"]:
                if a["name"] == account_name:
                    account_id = a["id"]
                    break
        if account_id is None:
            for a in self.settings["accounts"]:
                if a.get("enabled") and a.get("email"):
                    account_id = a["id"]
                    break
        provider = next((p for p in PROVIDERS if p["id"] == account_id), PROVIDERS[0])
        webbrowser.open(provider["url"])


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--background", action="store_true", help="run as the scheduled background notifier")
    args = parser.parse_args()
    if keyring is None:
        print("Missing dependency: keyring. Install requirements.txt first.")
        if not args.background:
            try:
                root = tk.Tk(); root.withdraw()
                messagebox.showerror("缺少组件", "请先安装依赖：pip install -r requirements.txt")
            except Exception:
                pass
        return 2
    if args.background:
        try:
            background_loop()
        except KeyboardInterrupt:
            pass
        return 0
    root = tk.Tk()
    app = MailBriefingApp(root)
    root.mainloop()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
