"""Write-operation local attachment coverage."""

import base64
import json
import os
import tempfile
import unittest
from contextlib import asynccontextmanager
from pathlib import Path
from unittest.mock import patch

import httpx

os.environ.setdefault("AZURE_TENANT_ID", "synthetic-tenant")
os.environ.setdefault("AZURE_CLIENT_ID", "synthetic-client")
os.environ.setdefault("AZURE_CLIENT_SECRET", "synthetic-secret")
os.environ.setdefault("OUTLOOK_USER_EMAIL", "mailbox@example.invalid")

import main


@asynccontextmanager
async def synthetic_graph(handler):
    client = httpx.AsyncClient(
        base_url="https://graph.microsoft.com/v1.0/users/mailbox@example.invalid",
        transport=httpx.MockTransport(handler),
    )
    previous_http = main._http
    main._http = client
    try:
        with patch.object(
            main,
            "_auth_headers",
            return_value={"Authorization": "Bearer synthetic-token"},
        ):
            yield
    finally:
        main._http = previous_http
        await client.aclose()


def request_json(request: httpx.Request) -> dict:
    return json.loads(request.content.decode("utf-8")) if request.content else {}


class WriteAttachmentTests(unittest.IsolatedAsyncioTestCase):
    async def test_send_email_without_attachments_uses_existing_sendmail_shape(self) -> None:
        seen: list[httpx.Request] = []

        async def handler(request: httpx.Request) -> httpx.Response:
            seen.append(request)
            return httpx.Response(202)

        async with synthetic_graph(handler):
            result = await main.send_email(
                to=["to@example.com"],
                subject="Subject",
                body="Body",
                cc=["cc@example.com"],
            )

        self.assertEqual(json.loads(result), {"status": "sent", "to": ["to@example.com"], "subject": "Subject"})
        self.assertEqual(len(seen), 1)
        self.assertEqual(seen[0].url.path, "/v1.0/users/mailbox@example.invalid/sendMail")
        payload = request_json(seen[0])
        self.assertNotIn("attachments", payload["message"])
        self.assertEqual(payload["message"]["body"], {"contentType": "Text", "content": "Body"})
        self.assertEqual(payload["message"]["ccRecipients"], [{"emailAddress": {"address": "cc@example.com"}}])

    async def test_send_email_encodes_multiple_local_attachments(self) -> None:
        seen_payloads: list[dict] = []
        with tempfile.TemporaryDirectory() as temp_dir:
            first = Path(temp_dir) / "alpha.txt"
            second = Path(temp_dir) / "report.bin"
            first.write_text("hello", encoding="utf-8")
            second.write_bytes(b"\x00\x01")

            async def handler(request: httpx.Request) -> httpx.Response:
                seen_payloads.append(request_json(request))
                return httpx.Response(202)

            async with synthetic_graph(handler):
                await main.send_email(
                    to=["to@example.com"],
                    subject="Attached",
                    body="See attached",
                    attachments=[str(first), str(second)],
                )

        attachments = seen_payloads[0]["message"]["attachments"]
        self.assertEqual([a["name"] for a in attachments], ["alpha.txt", "report.bin"])
        self.assertEqual(attachments[0]["@odata.type"], "#microsoft.graph.fileAttachment")
        self.assertEqual(attachments[0]["contentType"], "text/plain")
        self.assertEqual(attachments[0]["contentBytes"], base64.b64encode(b"hello").decode("ascii"))
        self.assertEqual(attachments[1]["contentType"], "application/octet-stream")
        self.assertEqual(attachments[1]["contentBytes"], base64.b64encode(b"\x00\x01").decode("ascii"))

    async def test_reply_email_without_attachments_uses_existing_reply_action(self) -> None:
        seen: list[httpx.Request] = []

        async def handler(request: httpx.Request) -> httpx.Response:
            seen.append(request)
            return httpx.Response(202)

        async with synthetic_graph(handler):
            result = await main.reply_email("message-id", "Thanks", reply_all=True)

        self.assertEqual(json.loads(result), {"status": "replied", "id": "message-id", "replyAll": True})
        self.assertEqual(len(seen), 1)
        self.assertEqual(seen[0].url.path, "/v1.0/users/mailbox@example.invalid/messages/message-id/replyAll")
        self.assertEqual(request_json(seen[0]), {"comment": "Thanks"})

    async def test_reply_email_with_attachments_creates_draft_adds_files_and_sends(self) -> None:
        seen: list[tuple[str, dict]] = []
        with tempfile.TemporaryDirectory() as temp_dir:
            first = Path(temp_dir) / "one.txt"
            second = Path(temp_dir) / "two.txt"
            first.write_text("one", encoding="utf-8")
            second.write_text("two", encoding="utf-8")

            async def handler(request: httpx.Request) -> httpx.Response:
                seen.append((request.url.path, request_json(request)))
                if request.url.path.endswith("/createReply"):
                    return httpx.Response(201, json={"id": "draft-reply"})
                if request.url.path.endswith("/attachments"):
                    return httpx.Response(201, json={"id": "attachment"})
                if request.url.path.endswith("/send"):
                    return httpx.Response(202)
                return httpx.Response(404, json={"error": {"message": "fixture miss"}})

            async with synthetic_graph(handler):
                await main.reply_email(
                    "message-id",
                    "HTML body",
                    body_type="HTML",
                    attachments=[str(first), str(second)],
                )

        self.assertEqual([path for path, _ in seen], [
            "/v1.0/users/mailbox@example.invalid/messages/message-id/createReply",
            "/v1.0/users/mailbox@example.invalid/messages/draft-reply/attachments",
            "/v1.0/users/mailbox@example.invalid/messages/draft-reply/attachments",
            "/v1.0/users/mailbox@example.invalid/messages/draft-reply/send",
        ])
        self.assertEqual(seen[0][1], {"message": {"body": {"contentType": "HTML", "content": "HTML body"}}})
        self.assertEqual([payload["name"] for _, payload in seen[1:3]], ["one.txt", "two.txt"])
        self.assertEqual(seen[1][1]["contentBytes"], base64.b64encode(b"one").decode("ascii"))
        self.assertEqual(seen[2][1]["contentBytes"], base64.b64encode(b"two").decode("ascii"))

    async def test_forward_email_without_attachments_uses_existing_forward_action(self) -> None:
        seen: list[httpx.Request] = []

        async def handler(request: httpx.Request) -> httpx.Response:
            seen.append(request)
            return httpx.Response(202)

        async with synthetic_graph(handler):
            result = await main.forward_email("message-id", ["to@example.com"], body="FYI")

        self.assertEqual(json.loads(result), {"status": "forwarded", "id": "message-id", "to": ["to@example.com"]})
        self.assertEqual(len(seen), 1)
        self.assertEqual(seen[0].url.path, "/v1.0/users/mailbox@example.invalid/messages/message-id/forward")
        self.assertEqual(request_json(seen[0]), {
            "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
            "comment": "FYI",
        })

    async def test_forward_email_with_attachments_creates_draft_adds_files_and_sends(self) -> None:
        seen: list[tuple[str, dict]] = []
        with tempfile.TemporaryDirectory() as temp_dir:
            first = Path(temp_dir) / "one.txt"
            second = Path(temp_dir) / "two.txt"
            first.write_text("one", encoding="utf-8")
            second.write_text("two", encoding="utf-8")

            async def handler(request: httpx.Request) -> httpx.Response:
                seen.append((request.url.path, request_json(request)))
                if request.url.path.endswith("/createForward"):
                    return httpx.Response(201, json={"id": "draft-forward"})
                if request.url.path.endswith("/attachments"):
                    return httpx.Response(201, json={"id": "attachment"})
                if request.url.path.endswith("/send"):
                    return httpx.Response(202)
                return httpx.Response(404, json={"error": {"message": "fixture miss"}})

            async with synthetic_graph(handler):
                await main.forward_email(
                    "message-id",
                    ["to@example.com"],
                    body="Forward body",
                    body_type="HTML",
                    attachments=[str(first), str(second)],
                )

        self.assertEqual([path for path, _ in seen], [
            "/v1.0/users/mailbox@example.invalid/messages/message-id/createForward",
            "/v1.0/users/mailbox@example.invalid/messages/draft-forward/attachments",
            "/v1.0/users/mailbox@example.invalid/messages/draft-forward/attachments",
            "/v1.0/users/mailbox@example.invalid/messages/draft-forward/send",
        ])
        self.assertEqual(seen[0][1], {
            "message": {
                "toRecipients": [{"emailAddress": {"address": "to@example.com"}}],
                "body": {"contentType": "HTML", "content": "Forward body"},
            }
        })
        self.assertEqual([payload["name"] for _, payload in seen[1:3]], ["one.txt", "two.txt"])

    async def test_invalid_attachment_path_fails_before_graph_write(self) -> None:
        calls = 0

        async def handler(request: httpx.Request) -> httpx.Response:
            nonlocal calls
            calls += 1
            return httpx.Response(500)

        missing = "/definitely/not/a/file.txt"
        async with synthetic_graph(handler):
            with self.assertRaisesRegex(ValueError, "Invalid attachment path"):
                await main.send_email(
                    to=["to@example.com"],
                    subject="Nope",
                    body="Body",
                    attachments=[missing],
                )
            with tempfile.TemporaryDirectory() as temp_dir:
                with self.assertRaisesRegex(ValueError, "not a regular file"):
                    await main.forward_email("message-id", ["to@example.com"], attachments=[temp_dir])

        self.assertEqual(calls, 0)


if __name__ == "__main__":
    unittest.main()
