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
import subprocess
import threading
import time
import webbrowser
from datetime import datetime, timedelta, timezone
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser
import uuid
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import keyring
except ImportError:
    keyring = None

APP_TITLE = "邮件助手"
UI_FONT = "PingFang SC" if sys.platform == "darwin" else ("Microsoft YaHei UI" if os.name == "nt" else "Noto Sans CJK SC")
SERVICE_NAME = "MailBriefingAssistant"
PROVIDERS = [
    {"id": "gmail", "name": "Gmail", "host": "imap.gmail.com", "port": 993,
     "url": "https://mail.google.com/mail/u/0/#inbox",
     "help": "填写 Google 应用专用密码；若账号不支持应用专用密码，当前版本暂不支持 Gmail OAuth 登录。"},
    {"id": "qq", "name": "QQ 邮箱", "host": "imap.qq.com", "port": 993,
     "url": "https://mail.qq.com/",
     "help": "先在邮箱设置中开启 IMAP 服务，并填写生成的授权码。"},
    {"id": "tsinghua", "name": "清华大学邮箱", "host": "mails.tsinghua.edu.cn", "port": 993,
     "url": "https://mails.tsinghua.edu.cn/",
     "help": "请填写清华邮箱设置中的客户端专用密码。"},
    {"id": "outlook", "name": "Outlook / Hotmail", "host": "outlook.office365.com", "port": 993,
     "url": "https://outlook.live.com/mail/",
     "help": "注意：部分 Outlook / Microsoft 365 账号只允许 OAuth2，当前版本尚不支持 OAuth2；仅在账号允许 IMAP 密码/应用密码登录时可用。"},
    {"id": "yahoo", "name": "Yahoo Mail", "host": "imap.mail.yahoo.com", "port": 993,
     "url": "https://mail.yahoo.com/",
     "help": "通常需要在 Yahoo 账号安全设置中生成应用专用密码。"},
    {"id": "icloud", "name": "iCloud Mail", "host": "imap.mail.me.com", "port": 993,
     "url": "https://www.icloud.com/mail/",
     "help": "使用 Apple 账号生成的 App 专用密码；IMAP 用户名通常为完整 iCloud 邮箱地址。"},
    {"id": "163", "name": "网易 163 邮箱", "host": "imap.163.com", "port": 993,
     "url": "https://mail.163.com/",
     "help": "先在邮箱设置中开启 IMAP/SMTP 服务，并使用客户端授权码。"},
    {"id": "126", "name": "网易 126 邮箱", "host": "imap.126.com", "port": 993,
     "url": "https://mail.126.com/",
     "help": "先在邮箱设置中开启 IMAP/SMTP 服务，并使用客户端授权码。"},
    {"id": "sina", "name": "新浪邮箱", "host": "imap.sina.com", "port": 993,
     "url": "https://mail.sina.com.cn/",
     "help": "请确认邮箱已开启 IMAP 服务，并按账号设置使用客户端授权码。"},
    {"id": "exmail", "name": "腾讯企业邮箱", "host": "imap.exmail.qq.com", "port": 993,
     "url": "https://exmail.qq.com/",
     "help": "请确认企业管理员允许 IMAP，并使用邮箱客户端密码/授权码。"},
    {"id": "aliyun", "name": "阿里云企业邮箱", "host": "imap.qiye.aliyun.com", "port": 993,
     "url": "https://qiye.aliyun.com/",
     "help": "企业邮箱服务器可能由管理员自定义；必要时请使用“其他 IMAP 邮箱”填写管理员提供的地址。"},
    {"id": "zoho", "name": "Zoho Mail", "host": "imap.zoho.com", "port": 993,
     "url": "https://mail.zoho.com/",
     "help": "服务器地址可能因账号的数据中心区域不同而变化；按 Zoho 账户设置核对 IMAP 地址。"},
    {"id": "fastmail", "name": "Fastmail", "host": "imap.fastmail.com", "port": 993,
     "url": "https://www.fastmail.com/",
     "help": "使用 Fastmail 的应用专用密码，不要使用网页登录密码。"},
]
DEFAULT_PROVIDER_IDS = {"gmail", "qq", "tsinghua"}
CATEGORY_ORDER = ["需尽快处理", "日程与会议", "课程与学术", "求职与实习", "账单与账户", "通知与订阅", "其他往来"]
CATEGORY_ALIASES = {
    "紧急/待处理": "需尽快处理",
    "订阅与通知": "通知与订阅",
    "个人与其他": "其他往来",
}
CATEGORY_RULES = {
    "需尽快处理": [
        "urgent", "action required", "asap", "due today", "deadline", "overdue",
        "请尽快", "尽快处理", "待处理", "待确认", "截止日期", "即将截止",
        "逾期", "需要回复", "请回复", "需回复", "待审批", "审核通过",
        "payment due", "verify your account", "account locked",
    ],
    "日程与会议": [
        "calendar", "meeting", "appointment", "invitation", "schedule", "webinar",
        "zoom", "teams meeting", "会议", "日程", "预约", "活动邀请", "面试安排",
        "报名确认", "时间变更", "地点变更", "开会", "日历",
    ],
    "课程与学术": [
        "清华", "北京大学", "课程", "教务", "研究生", "学术", "论文", "讲座",
        "seminar", "conference", "journal", "submission", "course", "university",
        "academic", "scholarship", "advisor", "导师", "答辩", "开题", "成绩",
    ],
    "求职与实习": [
        "招聘", "实习", "面试", "录用", "offer", "career", "recruit", "interview",
        "internship", "job opportunity", "快手", "美团", "简历", "岗位", "hr ",
        "申请进度", "候选人",
    ],
    "账单与账户": [
        "账单", "发票", "付款", "支付", "退款", "订单", "invoice", "payment",
        "receipt", "billing", "subscription renewal", "bank statement", "密码重置",
        "password reset", "登录提醒", "security alert", "verification code",
        "验证码", "安全提醒", "异常登录", "付款成功",
    ],
    "通知与订阅": [
        "unsubscribe", "newsletter", "邮件订阅", "每周摘要", "weekly digest",
        "no-reply", "noreply", "notification", "通知摘要", "产品更新",
        "product update", "digest", "自动回复", "系统通知",
    ],
}
PRIORITY_TERMS = {
    "高": ["urgent", "action required", "asap", "due today", "overdue", "deadline",
           "紧急", "尽快", "截止", "逾期", "请回复", "待审批", "账户锁定"],
    "中": ["meeting", "invitation", "interview", "payment", "calendar",
           "会议", "邀请", "面试", "付款", "需要确认", "请确认", "预约"],
}

def normalize_category(category: str) -> str:
    return CATEGORY_ALIASES.get(category, category if category in CATEGORY_ORDER else "其他往来")


def classify(mail: dict) -> str:
    haystack = " ".join(str(mail.get(key, "") or "") for key in
                        ("subject", "sender", "sender_address", "preview")).lower()
    for category in CATEGORY_ORDER[:-1]:
        if any(term.lower() in haystack for term in CATEGORY_RULES[category]):
            return category
    if any(x in haystack for x in ("edu.cn", "tsinghua.edu", "pku.edu", "course", "论文", "学术")):
        return "课程与学术"
    return "其他往来"


def classify_priority(mail: dict) -> str:
    """Transparent keyword-based triage; users should verify classification for critical mail."""
    haystack = " ".join(str(mail.get(key, "") or "") for key in
                        ("subject", "sender", "sender_address", "preview")).lower()
    if any(term.lower() in haystack for term in PRIORITY_TERMS["高"]):
        return "高"
    if any(term.lower() in haystack for term in PRIORITY_TERMS["中"]):
        return "中"
    return "普通"


def user_data_dir() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or (Path.home() / "AppData" / "Roaming"))
        path = base / "MailBriefingAssistant"
    elif sys.platform == "darwin":
        path = Path.home() / "Library" / "Application Support" / "MailAssistant"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or (Path.home() / ".config"))
        path = base / "MailBriefingAssistant"
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
    providers = [p for p in PROVIDERS if p["id"] in DEFAULT_PROVIDER_IDS]
    return {
        "daily_time": "08:00",
        "accounts": [{
            "id": p["id"], "provider_id": p["id"], "name": p["name"],
            "host": p["host"], "port": p["port"], "url": p["url"],
            "email": "", "enabled": True
        } for p in providers],
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
    if not isinstance(saved, dict):
        return result
    result["daily_time"] = saved.get("daily_time", result["daily_time"])
    saved_accounts = {
        str(account.get("id")): account
        for account in saved.get("accounts", [])
        if isinstance(account, dict) and account.get("id")
    }
    known_ids = {account["id"] for account in result["accounts"]}
    merged = []
    for account in result["accounts"]:
        old = saved_accounts.get(account["id"], {})
        merged.append({
            **account,
            **{key: old[key] for key in ("name", "host", "port", "url", "email", "enabled", "provider_id") if key in old},
            "provider_id": old.get("provider_id", old.get("id", account["id"])),
            "enabled": bool(old.get("enabled", account.get("enabled", True))),
        })
    # Preserve providers added in newer versions and user-defined IMAP accounts.
    for account_id, old in saved_accounts.items():
        if account_id in known_ids:
            continue
        provider_id = old.get("provider_id", account_id)
        provider = next((p for p in PROVIDERS if p["id"] == provider_id), None)
        if not isinstance(old.get("host"), str) or not old.get("host"):
            continue
        normalized = {
            "id": account_id,
            "provider_id": provider_id if provider else "custom",
            "name": old.get("name") or (provider["name"] if provider else "自定义邮箱"),
            "host": old.get("host"),
            "port": int(old.get("port") or 993),
            "url": old.get("url", provider["url"] if provider else ""),
            "email": old.get("email", ""),
            "enabled": bool(old.get("enabled", True)),
        }
        merged.append(normalized)
    result["accounts"] = merged
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
    """Store only local display metadata and a short body excerpt; never store full mail bodies or attachments."""
    date_value = mail.get("date")
    if isinstance(date_value, datetime):
        date_value = date_value.isoformat()
    return {
        "uid": mail.get("uid"),
        "account_id": mail.get("account_id"),
        "account_name": mail.get("account_name", ""),
        "provider_id": mail.get("provider_id", mail.get("account_id", "")),
        "subject": mail.get("subject", "（无主题）"),
        "sender": mail.get("sender", "未知发件人"),
        "category": normalize_category(mail.get("category", "其他往来")),
        "priority": mail.get("priority", classify_priority(mail)),
        "date": date_value,
        "preview": str(mail.get("preview", "") or "")[:220],
    }

def _restore_mail(record: dict) -> dict:
    mail = dict(record)
    date_value = mail.get("date")
    if isinstance(date_value, str):
        try:
            mail["date"] = datetime.fromisoformat(date_value)
        except ValueError:
            mail["date"] = None
    mail["category"] = normalize_category(mail.get("category", "其他往来"))
    mail.setdefault("priority", classify_priority(mail))
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


class VisibleHTMLParser(HTMLParser):
    """Extract visible text from HTML emails without executing or rendering HTML."""
    BLOCK_TAGS = {
        "address", "article", "br", "dd", "div", "dl", "dt", "footer",
        "h1", "h2", "h3", "h4", "header", "hr", "li", "main", "ol", "p",
        "section", "table", "tr", "ul",
    }
    HIDDEN_TAGS = {"script", "style", "head", "svg", "noscript"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.hidden_depth = 0

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag in self.HIDDEN_TAGS:
            self.hidden_depth += 1
        elif not self.hidden_depth and tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in self.HIDDEN_TAGS and self.hidden_depth:
            self.hidden_depth -= 1
        elif not self.hidden_depth and tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.hidden_depth:
            self.parts.append(data)

    def get_text(self) -> str:
        return re.sub(r"[ \t\u00a0]+", " ", re.sub(r"\n[ \t]*", "\n", "".join(self.parts))).strip()


def _part_text(part) -> str:
    try:
        value = part.get_content()
    except Exception:
        payload = part.get_payload(decode=True) or b""
        value = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
    if part.get_content_type() == "text/html":
        parser = VisibleHTMLParser()
        try:
            parser.feed(str(value))
            return parser.get_text()
        except Exception:
            return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", str(value))).strip()
    return str(value).strip()


def extract_body(message, limit: int = 450, preserve_lines: bool = False) -> str:
    """Prefer text/plain, otherwise extract visible text from HTML; ignore attachments."""
    plain_parts = []
    html_parts = []
    try:
        parts = message.walk() if message.is_multipart() else [message]
        for part in parts:
            if part.get_content_maintype() == "multipart":
                continue
            if part.get_content_disposition() == "attachment":
                continue
            content_type = part.get_content_type()
            if content_type not in ("text/plain", "text/html"):
                continue
            text = _part_text(part)
            if text:
                (plain_parts if content_type == "text/plain" else html_parts).append(text)
            if sum(len(item) for item in plain_parts) >= limit * 2:
                break
    except Exception:
        logging.debug("Could not extract message body", exc_info=True)
    chosen = plain_parts if plain_parts else html_parts
    if preserve_lines:
        text = "\n".join(chosen)
        text = re.sub(r"[ \t\u00a0]+", " ", text)
        text = re.sub(r"\n[ \t]*\n+", "\n\n", text).strip()
    else:
        text = re.sub(r"\s+", " ", " ".join(chosen)).strip()
    return text[:limit]


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
    mail = {
        "uid": uid, "account_id": account["id"], "account_name": account["name"],
        "provider_id": account.get("provider_id", account.get("id", "")),
        "account_email": account.get("email", ""), "subject": subject, "sender": sender,
        "sender_address": address or from_header, "date": received,
        "preview": extract_body(msg, 450),
    }
    mail["category"] = classify(mail)
    mail["priority"] = classify_priority(mail)
    return mail


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


def _fetch_bytes(conn, uid: int, section: str) -> bytes:
    status, data = conn.uid("fetch", str(uid), f"({section})")
    if status != "OK" or not data:
        raise RuntimeError(f"无法读取邮件 UID {uid}")
    chunks = [item[1] for item in data if isinstance(item, tuple) and isinstance(item[1], bytes)]
    if not chunks:
        raise RuntimeError(f"邮件 UID {uid} 未返回内容")
    return b"\r\n".join(chunks)


def fetch_raw(conn, uid: int, body_limit: int | None = 262144) -> bytes:
    """Fetch message headers plus a bounded body fragment; never mark it read."""
    header = _fetch_bytes(conn, uid, "BODY.PEEK[HEADER]")
    body_section = "BODY.PEEK[TEXT]" if body_limit is None else f"BODY.PEEK[TEXT]<0.{int(body_limit)}>"
    try:
        body = _fetch_bytes(conn, uid, body_section)
    except RuntimeError:
        body = b""
    return header.rstrip(b"\r\n") + b"\r\n\r\n" + body


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


def send_notification(mails: list[dict], errors: list[str], digest_title: str = "邮件助手") -> None:
    if not mails and not errors:
        logging.info("Daily check completed: no new messages")
        return
    counts = {category: sum(1 for mail in mails if mail["category"] == category) for category in CATEGORY_ORDER}
    count_line = " · ".join(f"{name} {count}" for name, count in counts.items() if count)
    details = []
    for mail in mails[:2]:
        account_name = mail.get("account_name", "邮箱")
        sender = mail.get("sender", "未知发件人")
        subject = mail.get("subject", "（无主题）")
        preview = str(mail.get("preview", "") or "").strip()
        line = f"{account_name} · {mail.get('category', '个人与其他')}｜{sender}: {subject}"
        if preview:
            line += "\n" + preview[:55]
        details.append(line)
    summary_parts = [part for part in (count_line, "\n".join(details)) if part]
    if len(mails) > 2:
        summary_parts.append(f"另有 {len(mails) - 2} 封，请打开邮件助手查看")
    if errors:
        summary_parts.append(f"{len(errors)} 个邮箱检查失败，请打开邮件助手查看")
    summary = "\n".join(summary_parts) or "没有发现新邮件"
    title = f"{digest_title}：发现 {len(mails)} 封新邮件" if mails else "邮件助手：部分邮箱检查失败"
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


def hydrate_missing_previews(records: list[dict], settings: dict) -> bool:
    """Backfill short previews for old locally cached messages using read-only IMAP."""
    missing = {}
    for record in records:
        account_id = str(record.get("account_id", ""))
        uid = record.get("uid")
        if record.get("preview") or not account_id or uid is None:
            continue
        try:
            uid = int(uid)
        except (TypeError, ValueError):
            continue
        missing.setdefault((account_id, uid), record)

    if not missing:
        return False

    accounts = {str(account.get("id")): account for account in settings.get("accounts", [])}
    changed = False
    for account_id in {key[0] for key in missing}:
        account = accounts.get(account_id)
        if not account or not account.get("enabled") or not account.get("email"):
            continue
        try:
            conn = imap_connect(account)
        except Exception as exc:
            logging.info("Could not backfill cached preview for %s: %s", account_id, type(exc).__name__)
            continue
        try:
            for (missing_account_id, uid), record in missing.items():
                if missing_account_id != account_id:
                    continue
                try:
                    parsed = parse_message(fetch_raw(conn, uid), uid, account)
                    preview = str(parsed.get("preview", "") or "")[:220]
                    if preview:
                        record["preview"] = preview
                        changed = True
                except Exception as exc:
                    logging.debug("Could not load cached message preview %s/%s: %s", account_id, uid, type(exc).__name__)
        finally:
            try:
                conn.logout()
            except Exception:
                pass

    # The same message may appear in both the digest and the pending queue.
    previews = {
        (str(record.get("account_id", "")), str(record.get("uid", ""))): record.get("preview", "")
        for record in records if record.get("preview")
    }
    for record in records:
        key = (str(record.get("account_id", "")), str(record.get("uid", "")))
        if not record.get("preview") and previews.get(key):
            record["preview"] = previews[key]
            changed = True
    return changed

class MailBriefingWidget:
    """Small always-on-top desktop card with collapsible email details."""

    BG = "#F5F7FA"
    PANEL = "#FFFFFF"
    ROW = "#FFFFFF"
    TEXT = "#1F2937"
    MUTED = "#667085"
    ACCENT = "#2563EB"
    BORDER = "#E4E7EC"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=self.BG)
        self.root.resizable(False, False)
        self.expanded = False
        self.drag_origin = None
        self.auto_collapse_job = None
        self.last_mails = []
        self.last_errors = []
        self.latest_digest = load_digest()
        initial_state = load_state()
        initial_pending = list(initial_state.get("pending_digest", []))
        self._pending_signature = tuple((item.get("account_id"), item.get("uid")) for item in initial_pending)
        if initial_pending:
            self.latest_digest = {
                "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                "mails": initial_pending,
                "errors": [],
            }
        self.last_mails = [_restore_mail(item) for item in self.latest_digest.get("mails", [])]
        self.last_errors = self.latest_digest.get("errors", [])
        self._build()
        self._update_summary()
        self._render_details()
        self._place(expanded=False)
        self._start_preview_hydration()
        self.root.after(1200, self._tick)

    def _button(self, parent, label, command, accent=False):
        widget = tk.Label(
            parent, text=label, bg=self.PANEL,
            fg=self.ACCENT if accent else self.MUTED,
            cursor="hand2",
            font=(UI_FONT, 9, "bold"),
            padx=7, pady=3
        )
        widget.bind("<Button-1>", lambda event: command())
        widget.bind("<Enter>", lambda event: widget.configure(fg=self.TEXT))
        widget.bind("<Leave>", lambda event: widget.configure(fg=self.ACCENT if accent else self.MUTED))
        return widget

    def _build(self):
        self.card = tk.Frame(self.root, bg=self.BG, highlightthickness=1, highlightbackground=self.BORDER)
        self.card.pack(fill="both", expand=True)
        header = tk.Frame(self.card, bg=self.BG)
        header.pack(fill="x", padx=12, pady=(9, 2))
        self.title_label = tk.Label(
            header, text="邮件助手", bg=self.BG, fg=self.TEXT,
            font=(UI_FONT, 11, "bold"), cursor="fleur"
        )
        self.title_label.pack(side="left")
        self.title_label.bind("<ButtonPress-1>", self._drag_start)
        self.title_label.bind("<B1-Motion>", self._drag_move)
        self.title_label.bind("<ButtonRelease-1>", self._drag_end)
        self.expand_button = self._button(header, "详情 ▾", self.toggle_details, accent=True)
        self.expand_button.pack(side="right")
        self.settings_button = self._button(header, "设置", self.open_settings)
        self.settings_button.pack(side="right", padx=(0, 4))
        self.min_button = self._button(header, "—", self.collapse_details)
        self.min_button.pack(side="right", padx=(0, 4))
        self.quit_button = self._button(header, "×", self.quit)
        self.quit_button.pack(side="right", padx=(0, 4))
        self.summary_label = tk.Label(
            self.card, text="", bg=self.BG, fg=self.MUTED,
            font=(UI_FONT, 9), justify="left", anchor="w",
            wraplength=400
        )
        self.summary_label.pack(fill="x", padx=13, pady=(1, 8))
        self.content = tk.Frame(self.card, bg=self.BG)
        self.section_label = tk.Label(
            self.content, text="新邮件详情  ·  点击邮件可打开对应邮箱",
            bg=self.BG, fg=self.TEXT, font=(UI_FONT, 9, "bold"), anchor="w"
        )
        self.section_label.pack(fill="x", padx=12, pady=(2, 7))
        list_holder = tk.Frame(self.content, bg=self.BG)
        list_holder.pack(fill="both", expand=True, padx=10)
        self.canvas = tk.Canvas(list_holder, bg=self.BG, highlightthickness=0, bd=0)
        self.scrollbar = tk.Scrollbar(list_holder, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.list_frame = tk.Frame(self.canvas, bg=self.BG)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.list_frame.bind("<Configure>", lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda event: self.canvas.itemconfigure(self.canvas_window, width=event.width))
        self.canvas.bind_all("<MouseWheel>", self._mousewheel)
        footer = tk.Frame(self.content, bg=self.BG)
        footer.pack(fill="x", padx=12, pady=(7, 10))
        tk.Label(
            footer, text="邮件仅在本机分类 · 不上传邮件内容",
            bg=self.BG, fg="#8291A8", font=(UI_FONT, 8)
        ).pack(side="left")
        self.open_mail_button = self._button(footer, "打开邮箱 ↗", self.open_webmail, accent=True)
        self.open_mail_button.pack(side="right")

    def _place(self, expanded=None):
        if expanded is None:
            expanded = self.expanded
        width = 440 if expanded else 360
        height = 560 if expanded else 104
        self.root.update_idletasks()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = max(8, screen_w - width - 22)
        y = max(8, screen_h - height - 58)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _drag_start(self, event):
        self.drag_origin = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())

    def _drag_move(self, event):
        if self.drag_origin:
            x = max(0, event.x_root - self.drag_origin[0])
            y = max(0, event.y_root - self.drag_origin[1])
            self.root.geometry(f"+{x}+{y}")

    def _drag_end(self, event):
        self.drag_origin = None

    def _mousewheel(self, event):
        if self.expanded:
            try:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except tk.TclError:
                pass

    def _cancel_auto_collapse(self):
        if self.auto_collapse_job is not None:
            try:
                self.root.after_cancel(self.auto_collapse_job)
            except tk.TclError:
                pass
            self.auto_collapse_job = None

    def _update_summary(self):
        settings = load_settings()
        due_time = settings.get("daily_time", "08:00")
        date_today = datetime.now().date().isoformat()
        generated = str(self.latest_digest.get("generated_at", ""))
        mail_count = len(self.last_mails)
        if generated[:10] == date_today:
            text = f"今日新邮件 {mail_count} 封  ·  点击详情查看分类与内容"
            if self.last_errors:
                text += f"\n{len(self.last_errors)} 个邮箱连接需检查"
        else:
            configured = sum(1 for a in settings["accounts"] if a.get("enabled") and a.get("email"))
            if configured:
                text = f"每天 {due_time} 检查  ·  已连接配置 {configured} 个邮箱"
            else:
                text = "点击“设置”添加邮箱，支持 Gmail、QQ 邮箱和清华邮箱"
        self.summary_label.configure(text=text, wraplength=410 if self.expanded else 332)

    def _render_details(self):
        for child in self.list_frame.winfo_children():
            child.destroy()

        counts = {
            category: sum(1 for mail in self.last_mails if mail.get("category") == category)
            for category in CATEGORY_ORDER
        }
        summary = "  ·  ".join(f"{category} {count}" for category, count in counts.items() if count)
        if summary:
            tk.Label(
                self.list_frame, text=summary, bg=self.BG, fg=self.MUTED,
                font=(UI_FONT, 8), justify="left",
                wraplength=400, anchor="w"
            ).pack(fill="x", padx=3, pady=(0, 8))

        mails = sorted(
            self.last_mails,
            key=lambda mail: mail.get("date") or datetime.min.replace(tzinfo=timezone.utc),
            reverse=True,
        )
        if not mails and not self.last_errors:
            tk.Label(
                self.list_frame,
                text="今天暂时没有新邮件。到设定时间后，这里会展示邮件来源、分类和正文摘要。",
                bg=self.ROW, fg=self.MUTED, font=(UI_FONT, 9),
                padx=12, pady=16, wraplength=400, justify="left",
            ).pack(fill="x", padx=2, pady=2)

        category_colors = {
            "紧急/待处理": "#B42318",
            "课程与学术": "#175CD3",
            "求职与实习": "#067647",
            "账单与账户": "#B54708",
            "订阅与通知": "#6941C6",
            "个人与其他": "#667085",
        }
        account_colors = {
            "gmail": ("#EAF2FF", "#1D4ED8"),
            "qq": ("#E8F7F2", "#067647"),
            "tsinghua": ("#FFF3E8", "#B54708"),
        }

        for mail in mails[:100]:
            row = tk.Frame(
                self.list_frame, bg=self.ROW, highlightthickness=1,
                highlightbackground=self.BORDER, cursor="hand2"
            )
            row.pack(fill="x", padx=2, pady=4)

            date_value = mail.get("date")
            if isinstance(date_value, datetime):
                date_text = date_value.astimezone().strftime("%m-%d %H:%M")
            elif date_value:
                date_text = str(date_value)[:16]
            else:
                date_text = "时间未知"

            top = tk.Frame(row, bg=self.ROW, cursor="hand2")
            top.pack(fill="x", padx=10, pady=(8, 3))

            account_id = str(mail.get("account_id", ""))
            badge_bg, badge_fg = account_colors.get(account_id, ("#EEF2F6", "#475467"))
            account_name = mail.get("account_name") or {
                "gmail": "Gmail",
                "qq": "QQ 邮箱",
                "tsinghua": "清华大学邮箱",
            }.get(account_id, "其他邮箱")
            account_label = tk.Label(
                top, text=account_name, bg=badge_bg, fg=badge_fg,
                font=(UI_FONT, 8, "bold"),
                padx=6, pady=2, cursor="hand2",
            )
            account_label.pack(side="left")

            category = mail.get("category", "个人与其他")
            category_label = tk.Label(
                top, text=category, bg=self.ROW,
                fg=category_colors.get(category, self.MUTED),
                font=(UI_FONT, 8, "bold"), cursor="hand2",
            )
            category_label.pack(side="left", padx=(8, 0))

            date_label = tk.Label(
                top, text=date_text, bg=self.ROW, fg=self.MUTED,
                font=(UI_FONT, 8), cursor="hand2",
            )
            date_label.pack(side="right")

            sender_label = tk.Label(
                row, text=mail.get("sender", "未知发件人"), bg=self.ROW, fg=self.MUTED,
                font=(UI_FONT, 9), anchor="w", justify="left", cursor="hand2",
            )
            sender_label.pack(fill="x", padx=10, pady=(0, 1))

            subject_label = tk.Label(
                row, text=mail.get("subject", "（无主题）"), bg=self.ROW, fg=self.TEXT,
                font=(UI_FONT, 10, "bold"), anchor="w", justify="left",
                wraplength=400, cursor="hand2",
            )
            subject_label.pack(fill="x", padx=10, pady=(1, 3))

            preview = str(mail.get("preview", "") or "").strip()
            if not preview:
                preview = "暂无正文摘要；点击此邮件可打开原邮箱查看。"
            preview_label = tk.Label(
                row, text=preview, bg=self.ROW, fg="#475467",
                font=(UI_FONT, 9), anchor="w", justify="left",
                wraplength=400, cursor="hand2",
            )
            preview_label.pack(fill="x", padx=10, pady=(0, 9))

            open_callback = lambda event, selected_account=account_id: self.open_webmail(selected_account)
            for part in (row, top, account_label, category_label, date_label, sender_label, subject_label, preview_label):
                part.bind("<Button-1>", open_callback)

        for error in self.last_errors:
            error_row = tk.Label(
                self.list_frame, text="检查提示  ·  " + error,
                bg="#FEF3F2", fg="#B42318",
                font=(UI_FONT, 8), anchor="w", justify="left",
                wraplength=400, padx=10, pady=8,
            )
            error_row.pack(fill="x", padx=2, pady=3)

        self.canvas.yview_moveto(0)

    def _set_digest(self, digest):
        self.latest_digest = digest if isinstance(digest, dict) else {}
        self.last_mails = [_restore_mail(item) for item in self.latest_digest.get("mails", [])]
        self.last_errors = self.latest_digest.get("errors", [])
        self._update_summary()
        self._render_details()

    def toggle_details(self):
        self._cancel_auto_collapse()
        if self.expanded:
            self.collapse_details()
        else:
            self.expand_details(auto=False)

    def expand_details(self, auto=False):
        self._cancel_auto_collapse()
        self.expanded = True
        self.expand_button.configure(text="收起 ▴")
        self.content.pack(fill="both", expand=True, padx=0, pady=(0, 1))
        self._update_summary()
        self._place(expanded=True)
        if auto:
            self.auto_collapse_job = self.root.after(25000, self.collapse_details)

    def collapse_details(self):
        self._cancel_auto_collapse()
        self.expanded = False
        self.expand_button.configure(text="详情 ▾")
        self.content.pack_forget()
        self._update_summary()
        self._place(expanded=False)

    def show_digest(self, digest, auto_collapse=True):
        self._set_digest(digest)
        self.root.deiconify()
        if self.last_mails or self.last_errors:
            self.expand_details(auto=auto_collapse)
        else:
            self.collapse_details()
            self.summary_label.configure(text="今日晨报已完成 · 暂无新邮件")
        try:
            self.root.attributes("-topmost", True)
            self.root.lift()
        except tk.TclError:
            pass

    def open_settings(self):
        self._cancel_auto_collapse()
        script = Path(__file__).resolve()
        kwargs = {}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        try:
            subprocess.Popen([sys.executable, str(script)], cwd=str(script.parent), **kwargs)
        except Exception as exc:
            logging.warning("Could not open settings window: %s", type(exc).__name__)

    def open_webmail(self, account_id=None):
        settings = load_settings()
        if account_id is None:
            account = next((a for a in settings["accounts"] if a.get("enabled") and a.get("email")), None)
        else:
            account = next((a for a in settings["accounts"] if a.get("id") == account_id), None)
        provider = next((p for p in PROVIDERS if account and p["id"] == account.get("id")), PROVIDERS[0])
        webbrowser.open(provider["url"])

    def quit(self):
        self._cancel_auto_collapse()
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def _start_preview_hydration(self):
        if getattr(self, "_preview_hydration_started", False):
            return
        self._preview_hydration_started = True

        def worker():
            digest = load_digest()
            state = load_state()
            digest_mails = digest.get("mails", []) if isinstance(digest.get("mails", []), list) else []
            pending_mails = state.get("pending_digest", []) if isinstance(state.get("pending_digest", []), list) else []
            records = digest_mails + pending_mails
            if not any(not record.get("preview") for record in records):
                return
            try:
                changed = hydrate_missing_previews(records, load_settings())
            except Exception:
                logging.debug("Cached preview refresh failed", exc_info=True)
                return
            if not changed:
                return
            try:
                digest["mails"] = [_stored_mail(item) for item in digest_mails]
                DIGEST_PATH.write_text(json.dumps(digest, ensure_ascii=False, indent=2), encoding="utf-8")
                state["pending_digest"] = [_stored_mail(item) for item in pending_mails]
                save_state(state)
            except OSError:
                logging.warning("Could not save refreshed local message previews", exc_info=True)
                return
            try:
                self.root.after(0, lambda: self._apply_preview_refresh(digest, state))
            except tk.TclError:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _apply_preview_refresh(self, digest, state):
        pending = state.get("pending_digest", [])
        if pending:
            display = {
                "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                "mails": pending,
                "errors": [],
            }
            self._pending_signature = tuple((item.get("account_id"), item.get("uid")) for item in pending)
        else:
            display = digest
        self._set_digest(display)



    def _tick(self):
        try:
            now = datetime.now()
            settings = load_settings()
            state = load_state()
            minute_key = now.strftime("%Y-%m-%d %H:%M")
            due = datetime.strptime(settings.get("daily_time", "08:00"), "%H:%M").time()
            configured = any(a.get("enabled") and a.get("email") for a in settings["accounts"])
            if configured and now.time() >= due and state.get("last_digest_date") != now.date().isoformat():
                # Avoid duplicate checks during the same minute if a prior attempt raises.
                if getattr(self, "_last_attempt_minute", "") != minute_key:
                    self._last_attempt_minute = minute_key
                    do_check(settings, state, notify=False, daily=True)
                    state = load_state()
                    state["last_digest_date"] = now.date().isoformat()
                    save_state(state)
                    digest = load_digest()
                    self._pending_signature = tuple()
                    self.show_digest(digest, auto_collapse=True)
            else:
                # Reflect mails collected by the separate settings window's manual check.
                pending = list(state.get("pending_digest", []))
                signature = tuple((item.get("account_id"), item.get("uid")) for item in pending)
                if pending and signature != self._pending_signature:
                    self._pending_signature = signature
                    digest = {
                        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                        "mails": pending,
                        "errors": [],
                    }
                    self.show_digest(digest, auto_collapse=True)
        except Exception as exc:
            logging.warning("Desktop widget scheduler error: %s", type(exc).__name__)
        try:
            self.root.after(30000, self._tick)
        except tk.TclError:
            pass

def background_loop() -> None:
    """Run the desktop mini-widget and its local daily scheduler."""
    logging.info("Desktop mail widget started")
    root = tk.Tk()
    MailBriefingWidget(root)
    root.mainloop()

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
        ttk.Label(wrapper, text="邮件助手", font=(UI_FONT, 18, "bold")).pack(anchor="w")
        ttk.Label(wrapper, text="汇总 Gmail、QQ 邮箱与清华邮箱的新邮件；本机分类，不上传邮件内容。", foreground="#555").pack(anchor="w", pady=(2, 12))
        self.notebook = ttk.Notebook(wrapper)
        self.notebook.pack(fill="both", expand=True)
        self.settings_tab = ttk.Frame(self.notebook, padding=12)
        self.inbox_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.settings_tab, text="邮箱与定时设置")
        self.notebook.add(self.inbox_tab, text="晨报记录")
        self.build_settings_tab()
        self.build_inbox_tab()
        bottom = ttk.Frame(wrapper)
        bottom.pack(fill="x", pady=(10, 0))
        self.status = tk.StringVar(value="请先配置邮箱并保存。")
        ttk.Label(bottom, textvariable=self.status).pack(side="left", fill="x", expand=True)
        ttk.Button(bottom, text="打开邮箱", command=self.open_webmail).pack(side="right", padx=(6, 0))
        ttk.Button(bottom, text="立即检查新邮件", command=self.check_now).pack(side="right")

    def build_settings_tab(self):
        ttk.Label(self.settings_tab, text="账号设置", font=(UI_FONT, 11, "bold")).pack(anchor="w")
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
        ttk.Label(toolbar, text="新邮件", font=(UI_FONT, 11, "bold")).pack(side="left")
        ttk.Button(toolbar, text="刷新列表", command=self.render_welcome).pack(side="right")

        columns = ("category", "account", "sender", "subject", "date")
        self.tree = ttk.Treeview(self.inbox_tab, columns=columns, show="headings", height=11)
        headings = {
            "category": "分类", "account": "邮箱来源", "sender": "发件人",
            "subject": "邮件主题", "date": "时间",
        }
        widths = {"category": 105, "account": 115, "sender": 140, "subject": 330, "date": 125}
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=widths[col], minwidth=75, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.show_selected_preview)
        self.tree.bind("<Double-1>", lambda event: self.open_webmail())

        detail = ttk.LabelFrame(self.inbox_tab, text="邮件摘要", padding=9)
        detail.pack(fill="x", pady=(8, 0))
        self.preview_var = tk.StringVar(value="选择一封邮件，在这里查看正文摘要。")
        ttk.Label(
            detail, textvariable=self.preview_var, justify="left",
            wraplength=820, anchor="w",
        ).pack(fill="x")

        self.summary = tk.StringVar(value="点击“立即检查新邮件”获取分类摘要。")
        ttk.Label(self.inbox_tab, textvariable=self.summary, wraplength=820, justify="left").pack(fill="x", pady=(8, 0))
        self.recent_mails = []

    def show_selected_preview(self, event=None):
        selected = self.tree.selection()
        if not selected:
            self.preview_var.set("选择一封邮件，在这里查看正文摘要。")
            return
        item_ids = list(self.tree.get_children())
        try:
            index = item_ids.index(selected[0])
            mail = self.recent_mails[index]
        except (ValueError, IndexError):
            self.preview_var.set("无法读取这封邮件的摘要，请刷新列表。")
            return

        preview = str(mail.get("preview", "") or "").strip()
        if not preview:
            preview = "这封邮件尚无缓存的正文摘要；可双击邮件打开原邮箱查看。"
        self.preview_var.set(
            f"邮箱：{mail.get('account_name', '未知邮箱')}    "
            f"发件人：{mail.get('sender', '未知发件人')}\n"
            f"主题：{mail.get('subject', '（无主题）')}\n\n"
            f"正文摘要：{preview}"
        )

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
            self.recent_mails = []
            self.preview_var.set("选择一封邮件，在这里查看正文摘要。")
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
        self.recent_mails = list(mails)
        for item in self.tree.get_children():
            self.tree.delete(item)
        counts = {c: 0 for c in CATEGORY_ORDER}
        for mail in self.recent_mails:
            counts[mail["category"]] = counts.get(mail["category"], 0) + 1
            date_value = mail.get("date")
            if isinstance(date_value, datetime):
                date_text = date_value.astimezone().strftime("%m-%d %H:%M")
            else:
                date_text = str(date_value)[:16] if date_value else "未知"
            self.tree.insert("", "end", values=(
                mail["category"], mail.get("account_name", "未知邮箱"),
                mail.get("sender", "未知发件人"), mail.get("subject", "（无主题）"), date_text,
            ))
        self.summary.set(title + "：" + ("；".join(f"{k} {v} 封" for k, v in counts.items() if v) or "0 封") +
                         ("\n检查失败：" + "；".join(errors) if errors else "") +
                         ("\n分类采用本地关键词规则，不会调用外部 AI 服务。"))
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.show_selected_preview()
        else:
            self.preview_var.set("没有可显示的邮件摘要。")
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
    parser.add_argument("--background", action="store_true", help="run the desktop widget at Windows logon")
    parser.add_argument("--widget", action="store_true", help="open the desktop widget")
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
    if args.background or args.widget:
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
