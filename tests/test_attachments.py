"""Synthetic fixture matrix for attachment listing and delivery."""

import asyncio
import hashlib
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


def attachment_handler(
    *,
    attachment_id: str,
    name: str,
    content: bytes,
    kind: str = "fileAttachment",
    content_type: str = "application/octet-stream",
    declared_size: int | None = None,
    is_inline: bool = False,
    stream: httpx.AsyncByteStream | None = None,
):
    async def handler(request: httpx.Request) -> httpx.Response:
        if request.url.path.endswith(f"/attachments/{attachment_id}/$value"):
            if stream is not None:
                return httpx.Response(200, stream=stream)
            return httpx.Response(200, content=content)
        if request.url.path.endswith(f"/attachments/{attachment_id}"):
            return httpx.Response(
                200,
                json={
                    "@odata.type": f"#microsoft.graph.{kind}",
                    "id": attachment_id,
                    "name": name,
                    "contentType": content_type,
                    "size": len(content) if declared_size is None else declared_size,
                    "isInline": is_inline,
                },
            )
        return httpx.Response(404, json={"error": {"message": "fixture miss"}})

    return handler


class DelayedStream(httpx.AsyncByteStream):
    async def __aiter__(self):
        yield b"partial"
        await asyncio.sleep(0.1)
        yield b"never-delivered"


class AttachmentListingTests(unittest.IsolatedAsyncioTestCase):
    async def test_paginated_listing_includes_multiple_attachment_kinds(self) -> None:
        requests: list[httpx.Request] = []

        async def handler(request: httpx.Request) -> httpx.Response:
            requests.append(request)
            if request.url.params.get("$skiptoken") == "synthetic-page-2":
                return httpx.Response(
                    200,
                    json={
                        "value": [
                            {
                                "@odata.type": "#microsoft.graph.referenceAttachment",
                                "id": "reference-1",
                                "name": "synthetic-link",
                                "contentType": "text/html",
                                "size": 0,
                            }
                        ]
                    },
                )
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "id": "file-1",
                            "name": "synthetic-statement.pdf",
                            "contentType": "application/pdf",
                            "size": 123,
                            "isInline": True,
                        },
                        {
                            "@odata.type": "#microsoft.graph.itemAttachment",
                            "id": "item-1",
                            "name": "synthetic-forwarded-message",
                            "contentType": "message/rfc822",
                            "size": 456,
                        },
                    ],
                    "@odata.nextLink": (
                        "https://graph.microsoft.com/v1.0/users/"
                        "mailbox@example.invalid/messages/message-page/attachments"
                        "?$skiptoken=synthetic-page-2"
                    ),
                },
            )

        async with synthetic_graph(handler):
            response = json.loads(await main.list_attachments("message-page"))

        self.assertEqual(len(requests), 2)
        self.assertEqual(
            [attachment["kind"] for attachment in response],
            ["fileAttachment", "itemAttachment", "referenceAttachment"],
        )
        self.assertTrue(response[0]["isInline"])

    async def test_zero_attachment_message_returns_empty_list(self) -> None:
        async def handler(request: httpx.Request) -> httpx.Response:
            return httpx.Response(200, json={"value": []})

        async with synthetic_graph(handler):
            response = json.loads(await main.list_attachments("message-empty"))

        self.assertEqual(response, [])

    async def test_untrusted_pagination_link_is_rejected(self) -> None:
        async def handler(request: httpx.Request) -> httpx.Response:
            return httpx.Response(
                200,
                json={
                    "value": [],
                    "@odata.nextLink": "https://example.invalid/credential-capture",
                },
            )

        async with synthetic_graph(handler):
            with self.assertRaisesRegex(RuntimeError, "untrusted nextLink"):
                await main.list_attachments("message-untrusted-page")


class AttachmentDownloadTests(unittest.IsolatedAsyncioTestCase):
    async def test_large_inline_binary_download_is_byte_faithful(self) -> None:
        minimum_size = 25 * 1024 * 1024
        content = (bytes(range(256)) * ((minimum_size // 256) + 1))[: minimum_size + 17]
        handler = attachment_handler(
            attachment_id="large-inline",
            name="synthetic-large.bin",
            content=content,
            is_inline=True,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "caller-selected" / "quarantine"
            async with synthetic_graph(handler):
                response = json.loads(
                    await main.download_attachment(
                        "message-large",
                        "large-inline",
                        str(destination),
                    )
                )

            self.assertGreaterEqual(response["size"], minimum_size)
            self.assertTrue(response["isInline"])
            self.assertEqual(Path(response["path"]).read_bytes(), content)
            self.assertEqual(response["sha256"], hashlib.sha256(content).hexdigest())

    async def test_filename_collision_and_symlink_are_not_overwritten(self) -> None:
        content = b"synthetic new statement"
        handler = attachment_handler(
            attachment_id="collision-file",
            name="statement.pdf",
            content=content,
            content_type="application/pdf",
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "downloads"
            destination.mkdir()
            (destination / "statement.pdf").write_bytes(b"existing statement")
            outside = Path(temporary_root) / "outside.pdf"
            outside.write_bytes(b"outside sentinel")
            (destination / "statement (1).pdf").symlink_to(outside)

            async with synthetic_graph(handler):
                response = json.loads(
                    await main.download_attachment(
                        "message-collision",
                        "collision-file",
                        str(destination),
                    )
                )

            self.assertEqual(response["filename"], "statement (2).pdf")
            self.assertEqual(
                (destination / "statement.pdf").read_bytes(), b"existing statement"
            )
            self.assertEqual(outside.read_bytes(), b"outside sentinel")
            self.assertEqual(Path(response["path"]).read_bytes(), content)

    async def test_hostile_traversal_filename_stays_in_created_destination(
        self,
    ) -> None:
        content = b"synthetic hostile-name fixture"
        handler = attachment_handler(
            attachment_id="hostile-file",
            name="../../synthetic/../../escape.bin",
            content=content,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "new" / "quarantine"
            async with synthetic_graph(handler):
                response = json.loads(
                    await main.download_attachment(
                        "message-hostile",
                        "hostile-file",
                        str(destination),
                    )
                )

            delivered = Path(response["path"])
            self.assertEqual(delivered.parent, destination.resolve())
            self.assertNotIn("..", response["filename"])
            self.assertNotIn("/", response["filename"])
            self.assertEqual(delivered.read_bytes(), content)
            self.assertFalse((Path(temporary_root) / "escape.bin").exists())

    async def test_item_and_reference_downloads_raise_typed_error(self) -> None:
        for kind in ("itemAttachment", "referenceAttachment"):
            with self.subTest(kind=kind):
                handler = attachment_handler(
                    attachment_id=f"unsupported-{kind}",
                    name=f"synthetic-{kind}",
                    content=b"",
                    kind=kind,
                )
                with tempfile.TemporaryDirectory() as destination:
                    async with synthetic_graph(handler):
                        with self.assertRaises(
                            main.UnsupportedAttachmentKindError
                        ) as raised:
                            await main.download_attachment(
                                "message-unsupported",
                                f"unsupported-{kind}",
                                destination,
                            )
                    self.assertEqual(
                        raised.exception.code, "unsupported_attachment_kind"
                    )
                    self.assertEqual(raised.exception.kind, kind)
                    self.assertIn("unsupported_attachment_kind", str(raised.exception))

    async def test_destination_directory_symlink_is_rejected(self) -> None:
        content = b"synthetic symlink-destination fixture"
        handler = attachment_handler(
            attachment_id="destination-symlink",
            name="synthetic-safe-name.bin",
            content=content,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            actual_destination = Path(temporary_root) / "actual"
            actual_destination.mkdir()
            linked_destination = Path(temporary_root) / "linked"
            linked_destination.symlink_to(actual_destination, target_is_directory=True)

            async with synthetic_graph(handler):
                with self.assertRaises(main.UnsafeAttachmentDestinationError):
                    await main.download_attachment(
                        "message-destination-symlink",
                        "destination-symlink",
                        str(linked_destination),
                    )
            self.assertEqual(list(actual_destination.iterdir()), [])

    async def test_size_cap_removes_partial_file(self) -> None:
        handler = attachment_handler(
            attachment_id="over-limit",
            name="synthetic-over-limit.bin",
            content=b"0123456789",
            declared_size=0,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "downloads"
            async with synthetic_graph(handler):
                with self.assertRaises(main.AttachmentTooLargeError):
                    await main.download_attachment(
                        "message-over-limit",
                        "over-limit",
                        str(destination),
                        max_bytes=5,
                    )
            self.assertEqual(list(destination.iterdir()), [])

    async def test_total_timeout_removes_partial_file(self) -> None:
        handler = attachment_handler(
            attachment_id="slow-file",
            name="synthetic-slow.bin",
            content=b"",
            declared_size=0,
            stream=DelayedStream(),
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "downloads"
            async with synthetic_graph(handler):
                with self.assertRaises(TimeoutError):
                    await main.download_attachment(
                        "message-slow",
                        "slow-file",
                        str(destination),
                        timeout_seconds=0.01,
                    )
            self.assertEqual(list(destination.iterdir()), [])

    async def test_mime_vs_decoded_size_delta_is_not_treated_as_corruption(
        self,
    ) -> None:
        """Graph's declared MIME size legitimately exceeds the decoded bytes.

        The delivered bytes (an unrecognized format) carry no structural claim,
        so they must be accepted verbatim even though the declared size differs.
        Delivered-vs-declared byte counts must never be compared for equality.
        """
        content = b"decoded octet payload"
        handler = attachment_handler(
            attachment_id="mime-delta-file",
            name="synthetic-decoded.bin",
            content=content,
            declared_size=len(content) + 128,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "downloads"
            async with synthetic_graph(handler):
                response = json.loads(
                    await main.download_attachment(
                        "message-mime-delta",
                        "mime-delta-file",
                        str(destination),
                    )
                )
            self.assertEqual(Path(response["path"]).read_bytes(), content)
            self.assertEqual(response["size"], len(content))


class ReadOnlyScopeTests(unittest.TestCase):
    def test_default_attachment_guards_match_contract(self) -> None:
        self.assertEqual(main.DEFAULT_ATTACHMENT_MAX_BYTES, 100 * 1024 * 1024)
        self.assertEqual(main.DEFAULT_ATTACHMENT_TIMEOUT_SECONDS, 30.0)

    def test_requested_mail_permissions_are_read_only(self) -> None:
        self.assertLessEqual(
            main.REQUESTED_GRAPH_PERMISSIONS,
            main.READ_ONLY_GRAPH_PERMISSIONS,
        )
        self.assertLessEqual(
            main.REQUESTED_MAIL_PERMISSIONS,
            main.READ_ONLY_MAIL_PERMISSIONS,
        )
        for permission in main.REQUESTED_GRAPH_PERMISSIONS:
            self.assertNotIn("Send", permission)
            self.assertNotIn("ReadWrite", permission)


if __name__ == "__main__":
    unittest.main()
