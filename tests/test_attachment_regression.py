"""Regression coverage for byte-faithful attachment delivery."""

import base64
import hashlib
import json
import os
import tempfile
import unittest
from datetime import timedelta
from pathlib import Path

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
        attachment_bytes = (
            b"PK\x03\x04"
            + (b"\x00" * (256 * 1024))
            + bytes(range(256))
            + b"synthetic fixture only"
        )
        expected_sha256 = hashlib.sha256(attachment_bytes).hexdigest()

        async def graph_fixture(request: httpx.Request) -> httpx.Response:
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
            async with create_connected_server_and_client_session(
                main.mcp,
                read_timeout_seconds=timedelta(seconds=5),
            ) as session:
                lifespan_http = main._http
                main._http = fixture_http
                try:
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
            self.assertFalse(result.isError, response_text)
            response = json.loads(response_text)
            delivered_path = Path(response["path"])

            self.assertEqual(delivered_path.read_bytes(), attachment_bytes)
            self.assertEqual(response["sha256"], expected_sha256)
            self.assertEqual(response["size"], len(attachment_bytes))
            self.assertNotIn("contentBytes", response)


if __name__ == "__main__":
    unittest.main()
