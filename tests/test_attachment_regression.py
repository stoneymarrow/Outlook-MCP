"""Regression coverage for byte-faithful attachment delivery."""

import base64
import hashlib
import io
import json
import os
import tempfile
import unittest
import zipfile
from datetime import timedelta
from pathlib import Path
from unittest.mock import patch

import httpx
from mcp.shared.memory import create_connected_server_and_client_session

os.environ.setdefault("AZURE_TENANT_ID", "synthetic-tenant")
os.environ.setdefault("AZURE_CLIENT_ID", "synthetic-client")
os.environ.setdefault("AZURE_CLIENT_SECRET", "synthetic-secret")
os.environ.setdefault("OUTLOOK_USER_EMAIL", "mailbox@example.invalid")

import main


class AttachmentTransportRegressionTest(unittest.IsolatedAsyncioTestCase):
    async def test_repeated_binary_runs_remain_byte_identical(self) -> None:
        """Binary bytes must never cross the MCP response as base64 text."""
        # A real ZIP container (like the xlsx it mimics): high-byte binary body
        # plus a valid end-of-central-directory record, so it survives both the
        # base64-transport check and the structural integrity gate.
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as archive:
            archive.writestr("sheet1.bin", bytes(range(256)) * 1024)
            archive.writestr("notes.txt", b"synthetic fixture only")
        attachment_bytes = zip_buffer.getvalue()
        expected_sha256 = hashlib.sha256(attachment_bytes).hexdigest()

        async def graph_fixture(request: httpx.Request) -> httpx.Response:
            if request.url.path.endswith("/message-binary/attachments"):
                return httpx.Response(
                    200,
                    json={
                        "value": [
                            {
                                "@odata.type": "#microsoft.graph.fileAttachment",
                                "id": "attachment-binary",
                                "name": "synthetic-register.xlsx",
                                "contentType": (
                                    "application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet"
                                ),
                                "size": len(attachment_bytes),
                                "isInline": False,
                            }
                        ]
                    },
                )
            if request.url.path.endswith("/attachments/attachment-binary"):
                return httpx.Response(
                    200,
                    json={
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "id": "attachment-binary",
                        "name": "synthetic-register.xlsx",
                        "contentType": (
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        "size": len(attachment_bytes),
                        "isInline": False,
                        "contentBytes": base64.b64encode(attachment_bytes).decode(),
                    },
                )
            if request.url.path.endswith("/attachments/attachment-binary/$value"):
                return httpx.Response(200, content=attachment_bytes)
            return httpx.Response(404, json={"error": {"message": "fixture miss"}})

        fixture_http = httpx.AsyncClient(
            base_url="https://graph.microsoft.com/v1.0/users/mailbox@example.invalid",
            transport=httpx.MockTransport(graph_fixture),
        )
        with tempfile.TemporaryDirectory() as destination:
            with patch.object(
                main,
                "_auth_headers",
                return_value={"Authorization": "Bearer synthetic-token"},
            ):
                async with create_connected_server_and_client_session(
                    main.mcp,
                    read_timeout_seconds=timedelta(seconds=5),
                ) as session:
                    lifespan_http = main._http
                    main._http = fixture_http
                    try:
                        listing_result = await session.call_tool(
                            "list_attachments",
                            {"message_id": "message-binary"},
                        )
                        result = await session.call_tool(
                            "download_attachment",
                            {
                                "message_id": "message-binary",
                                "attachment_id": "attachment-binary",
                                "destination_directory": destination,
                            },
                        )
                    finally:
                        main._http = lifespan_http

            await fixture_http.aclose()

            response_text = "".join(
                block.text for block in result.content if hasattr(block, "text")
            )
            listing_text = "".join(
                block.text for block in listing_result.content if hasattr(block, "text")
            )
            self.assertFalse(listing_result.isError, listing_text)
            self.assertFalse(result.isError, response_text)
            listing = json.loads(listing_text)
            response = json.loads(response_text)
            delivered_path = Path(response["path"])

            self.assertEqual(listing[0]["kind"], "fileAttachment")
            self.assertEqual(delivered_path.read_bytes(), attachment_bytes)
            self.assertEqual(response["originalFilename"], "synthetic-register.xlsx")
            self.assertEqual(response["filename"], "synthetic-register.xlsx")
            self.assertEqual(
                response["contentType"],
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.assertEqual(response["sha256"], expected_sha256)
            self.assertEqual(response["size"], len(attachment_bytes))
            self.assertEqual(response["kind"], "fileAttachment")
            self.assertNotIn("contentBytes", response)


if __name__ == "__main__":
    unittest.main()
