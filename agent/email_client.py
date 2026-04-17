"""Microsoft 365 (Graph API) email client.

Uses the client-credentials flow (application permissions) so the agent
can run unattended on a VPS.
"""
from __future__ import annotations
import base64
import logging
import mimetypes
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, List, Optional

import requests
from msal import ConfidentialClientApplication
from tenacity import retry, stop_after_attempt, wait_exponential

from .config import Settings

log = logging.getLogger("ice_quote_agent")

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPE = ["https://graph.microsoft.com/.default"]


@dataclass
class EmailAttachment:
    name: str
    content_bytes: bytes
    content_type: str


@dataclass
class IncomingEmail:
    id: str
    subject: str
    sender: str
    body_text: str
    attachments: List[EmailAttachment] = field(default_factory=list)


class GraphClient:
    def __init__(self, settings: Settings):
        self.s = settings
        self._app = ConfidentialClientApplication(
            client_id=settings.client_id,
            client_credential=settings.client_secret,
            authority=f"https://login.microsoftonline.com/{settings.tenant_id}",
        )
        self._token: Optional[str] = None

    def _acquire_token(self) -> str:
        result = self._app.acquire_token_for_client(scopes=SCOPE)
        if "access_token" not in result:
            raise RuntimeError(f"MSAL token acquisition failed: {result}")
        return result["access_token"]

    def _headers(self) -> dict:
        if not self._token:
            self._token = self._acquire_token()
        return {
            "Authorization": f"Bearer {self._token}",
            "Content-Type": "application/json",
        }

    def _mailbox_url(self) -> str:
        return f"{GRAPH}/users/{self.s.mailbox}"

    @retry(stop=stop_after_attempt(3),
           wait=wait_exponential(multiplier=2, min=2, max=30))
    def fetch_unread(self, top: int = 10) -> List[IncomingEmail]:
        url = (f"{self._mailbox_url()}/mailFolders/Inbox/messages"
               f"?$filter=isRead eq false&$top={top}"
               f"&$select=id,subject,from,body,hasAttachments")
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        items = r.json().get("value", [])
        out: List[IncomingEmail] = []
        for m in items:
            sender = (m.get("from", {}) or {}).get("emailAddress", {}).get(
                "address", "")
            body = (m.get("body", {}) or {}).get("content", "") or ""
            if (m.get("body", {}) or {}).get("contentType") == "html":
                body = _strip_html(body)
            email = IncomingEmail(
                id=m["id"],
                subject=m.get("subject", ""),
                sender=sender,
                body_text=body,
                attachments=[],
            )
            if m.get("hasAttachments"):
                email.attachments = self._fetch_attachments(m["id"])
            out.append(email)
        return out

    @retry(stop=stop_after_attempt(3),
           wait=wait_exponential(multiplier=2, min=2, max=30))
    def _fetch_attachments(self, message_id: str) -> List[EmailAttachment]:
        url = f"{self._mailbox_url()}/messages/{message_id}/attachments"
        r = requests.get(url, headers=self._headers(), timeout=60)
        r.raise_for_status()
        out = []
        for att in r.json().get("value", []):
            if att.get("@odata.type") == "#microsoft.graph.fileAttachment":
                out.append(EmailAttachment(
                    name=att["name"],
                    content_bytes=base64.b64decode(att["contentBytes"]),
                    content_type=att.get("contentType") or
                                 mimetypes.guess_type(att["name"])[0] or
                                 "application/octet-stream",
                ))
        return out

    @retry(stop=stop_after_attempt(3),
           wait=wait_exponential(multiplier=2, min=2, max=30))
    def send_reply(
        self,
        to: str,
        cc: Optional[str],
        subject: str,
        body_text: str,
        attachment_path: Optional[Path] = None,
        reply_to_message_id: Optional[str] = None,
    ) -> None:
        attachments_payload = []
        if attachment_path is not None and attachment_path.exists():
            with open(attachment_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")
            attachments_payload.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": attachment_path.name,
                "contentBytes": b64,
                "contentType": mimetypes.guess_type(attachment_path.name)[0] or
                               "application/octet-stream",
            })

        recipients = [{"emailAddress": {"address": to}}]
        cc_list = [{"emailAddress": {"address": cc}}] if cc else []

        message = {
            "subject": subject,
            "body": {"contentType": "Text", "content": body_text},
            "toRecipients": recipients,
            "ccRecipients": cc_list,
            "attachments": attachments_payload,
        }

        if reply_to_message_id:
            create_url = (f"{self._mailbox_url()}/messages/"
                          f"{reply_to_message_id}/createReply")
            r = requests.post(create_url, headers=self._headers(), timeout=30)
            r.raise_for_status()
            draft_id = r.json()["id"]
            patch_url = f"{self._mailbox_url()}/messages/{draft_id}"
            r = requests.patch(patch_url, headers=self._headers(),
                               json=message, timeout=60)
            r.raise_for_status()
            send_url = f"{self._mailbox_url()}/messages/{draft_id}/send"
            r = requests.post(send_url, headers=self._headers(), timeout=60)
            r.raise_for_status()
        else:
            send_url = f"{self._mailbox_url()}/sendMail"
            r = requests.post(send_url, headers=self._headers(),
                              json={"message": message,
                                    "saveToSentItems": True},
                              timeout=60)
            r.raise_for_status()

    @retry(stop=stop_after_attempt(3),
           wait=wait_exponential(multiplier=2, min=2, max=30))
    def mark_read(self, message_id: str) -> None:
        url = f"{self._mailbox_url()}/messages/{message_id}"
        r = requests.patch(url, headers=self._headers(),
                           json={"isRead": True}, timeout=30)
        r.raise_for_status()

    def move_to_folder(self, message_id: str, folder_display_name: str) -> None:
        folder_id = self._get_or_create_folder(folder_display_name)
        if folder_id is None:
            return
        url = f"{self._mailbox_url()}/messages/{message_id}/move"
        r = requests.post(url, headers=self._headers(),
                          json={"destinationId": folder_id}, timeout=30)
        if not r.ok:
            log.warning("Move to %s failed: %s", folder_display_name, r.text)

    def _get_or_create_folder(self, name: str) -> Optional[str]:
        list_url = f"{self._mailbox_url()}/mailFolders?$top=100"
        r = requests.get(list_url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        for f in r.json().get("value", []):
            if f.get("displayName", "").lower() == name.lower():
                return f["id"]
        create_url = f"{self._mailbox_url()}/mailFolders"
        r = requests.post(create_url, headers=self._headers(),
                          json={"displayName": name}, timeout=30)
        if r.ok:
            return r.json()["id"]
        log.warning("Could not create folder %s: %s", name, r.text)
        return None


def _strip_html(html: str) -> str:
    from html.parser import HTMLParser

    class _Text(HTMLParser):
        def __init__(self):
            super().__init__()
            self.buf = []
            self._skip = False

        def handle_starttag(self, tag, attrs):
            if tag in ("script", "style"):
                self._skip = True
            if tag in ("br", "p", "li", "tr"):
                self.buf.append("\n")

        def handle_endtag(self, tag):
            if tag in ("script", "style"):
                self._skip = False

        def handle_data(self, data):
            if not self._skip:
                self.buf.append(data)

    p = _Text()
    p.feed(html)
    return "".join(p.buf)
