"""Regression coverage for delivered-byte integrity validation.

Graph's ``size`` field reports the MIME/base64-encoded size of an attachment,
while the raw ``$value`` endpoint returns the decoded bytes. Comparing the two
is wrong by design: a valid decoded payload is legitimately smaller than the
declared encoded size. The download path must validate the delivered bytes on
their own terms (hash + structural integrity), never against Graph's MIME size,
and must never delete a download that fails validation.
"""

import hashlib
import io
import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path

os.environ.setdefault("AZURE_TENANT_ID", "synthetic-tenant")
os.environ.setdefault("AZURE_CLIENT_ID", "synthetic-client")
os.environ.setdefault("AZURE_CLIENT_SECRET", "synthetic-secret")
os.environ.setdefault("OUTLOOK_USER_EMAIL", "mailbox@example.invalid")

import main

from tests.test_attachments import attachment_handler, synthetic_graph


def make_pdf(total_size: int) -> bytes:
    """Return valid-looking PDF bytes of exactly ``total_size`` length.

    Real PDFs open with ``%PDF-`` and close with a ``%%EOF`` trailer; the body
    in between is padded with a comment so the fixture reaches an exact length
    without disturbing either structural marker.
    """
    header = b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    trailer = b"\n%%EOF\n"
    filler_prefix = b"% "
    overhead = len(header) + len(filler_prefix) + len(trailer)
    if total_size < overhead:
        raise ValueError("total_size too small for a structural PDF fixture")
    filler = b"a" * (total_size - overhead)
    payload = header + filler_prefix + filler + trailer
    assert len(payload) == total_size, (len(payload), total_size)
    return payload


def make_zip() -> bytes:
    """Return a real, minimal ZIP container (carries an end-of-central-directory)."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("sheet.xml", b"<sheet/>synthetic fixture only")
    return buffer.getvalue()


class DeliveredByteIntegrityTests(unittest.IsolatedAsyncioTestCase):
    async def test_observed_mime_vs_decoded_delta_is_accepted(self) -> None:
        """The exact June live-run pair (declared 66388 / delivered 62553).

        Graph declared the MIME/base64 size (66388); the raw endpoint delivered
        the decoded PDF (62553). The old gate rejected this valid payload as a
        size mismatch. The new gate must accept it.
        """
        decoded = make_pdf(62553)
        self.assertEqual(len(decoded), 62553)
        handler = attachment_handler(
            attachment_id="mime-vs-decoded",
            name="portfolio-appraisal.pdf",
            content=decoded,
            content_type="application/pdf",
            declared_size=66388,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "quarantine"
            async with synthetic_graph(handler):
                response = json.loads(
                    await main.download_attachment(
                        "message-mime-vs-decoded",
                        "mime-vs-decoded",
                        str(destination),
                    )
                )

            delivered = Path(response["path"])
            self.assertEqual(delivered.read_bytes(), decoded)
            self.assertEqual(response["size"], 62553)
            self.assertEqual(response["sha256"], hashlib.sha256(decoded).hexdigest())

    async def test_second_observed_pair_is_accepted(self) -> None:
        """The second June pair (declared 66782 / delivered 62947)."""
        decoded = make_pdf(62947)
        handler = attachment_handler(
            attachment_id="mime-vs-decoded-2",
            name="transaction-statement.pdf",
            content=decoded,
            content_type="application/pdf",
            declared_size=66782,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "quarantine"
            async with synthetic_graph(handler):
                response = json.loads(
                    await main.download_attachment(
                        "message-mime-vs-decoded-2",
                        "mime-vs-decoded-2",
                        str(destination),
                    )
                )
            self.assertEqual(Path(response["path"]).read_bytes(), decoded)
            self.assertEqual(response["size"], 62947)

    async def test_valid_zip_container_with_size_delta_is_accepted(self) -> None:
        """A real ZIP (xlsx) delivered smaller than its declared MIME size passes."""
        decoded = make_zip()
        handler = attachment_handler(
            attachment_id="zip-delta",
            name="appraisal.xlsx",
            content=decoded,
            content_type=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            declared_size=len(decoded) + 4096,
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "quarantine"
            async with synthetic_graph(handler):
                response = json.loads(
                    await main.download_attachment(
                        "message-zip-delta",
                        "zip-delta",
                        str(destination),
                    )
                )
            self.assertEqual(Path(response["path"]).read_bytes(), decoded)

    async def test_truncated_pdf_fails_and_file_is_preserved(self) -> None:
        """A genuinely truncated PDF (lost %%EOF) must fail without deletion."""
        whole = make_pdf(20000)
        truncated = whole[:12000]  # drops the %%EOF trailer
        handler = attachment_handler(
            attachment_id="truncated-pdf",
            name="truncated.pdf",
            content=truncated,
            content_type="application/pdf",
            declared_size=len(whole),
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "quarantine"
            async with synthetic_graph(handler):
                with self.assertRaises(main.AttachmentIntegrityError) as raised:
                    await main.download_attachment(
                        "message-truncated-pdf",
                        "truncated-pdf",
                        str(destination),
                    )

            error = raised.exception
            self.assertEqual(error.code, "attachment_integrity_failed")
            # The bytes are preserved for inspection, never destroyed.
            preserved = Path(error.path)
            self.assertTrue(preserved.exists())
            self.assertEqual(preserved.read_bytes(), truncated)
            self.assertEqual(
                sorted(p.name for p in destination.iterdir()), [preserved.name]
            )

    async def test_truncated_zip_fails_and_file_is_preserved(self) -> None:
        """A truncated ZIP (lost end-of-central-directory) must fail without deletion."""
        whole = make_zip()
        truncated = whole[: len(whole) // 2]
        handler = attachment_handler(
            attachment_id="truncated-zip",
            name="truncated.xlsx",
            content=truncated,
            content_type=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            declared_size=len(whole),
        )

        with tempfile.TemporaryDirectory() as temporary_root:
            destination = Path(temporary_root) / "quarantine"
            async with synthetic_graph(handler):
                with self.assertRaises(main.AttachmentIntegrityError) as raised:
                    await main.download_attachment(
                        "message-truncated-zip",
                        "truncated-zip",
                        str(destination),
                    )
            preserved = Path(raised.exception.path)
            self.assertTrue(preserved.exists())
            self.assertEqual(preserved.read_bytes(), truncated)


if __name__ == "__main__":
    unittest.main()
