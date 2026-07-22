"""Regression coverage for the stdio-shutdown self-abort.

The server drives stdin/stdout through an anyio worker thread. If interpreter
finalization runs while that worker is still blocked mid-``readline`` (holding
the buffered-stream lock), CPython's ``_enter_buffered_busy`` fatal-errors and
``abort()``s the process (SIGABRT), producing a macOS crash report. The fix
(``main._run_stdio`` exiting via ``os._exit`` only after transport + lifespan
cleanup) must make every disconnect path exit cleanly instead.

These probes are fully contained: the server subprocess runs with synthetic
environment values, speaks only the MCP ``initialize`` / ``tools/list``
handshake (no Graph calls, no real credentials, ``.env`` unread), and is driven
over an in-process pipe. Each disconnect is repeated many times because the
underlying fault is an intermittent race - a single clean cycle proves nothing.
"""

import json
import os
import select
import signal
import subprocess
import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
MAIN_PY = REPO_ROOT / "main.py"

# Enough isolated cycles that an intermittent finalization race would surface.
CYCLES = 15
TIMEOUT_SECONDS = 15.0

SYNTHETIC_ENV = {
    "AZURE_TENANT_ID": "synthetic-tenant",
    "AZURE_CLIENT_ID": "synthetic-client",
    "AZURE_CLIENT_SECRET": "synthetic-secret",
    "OUTLOOK_USER_EMAIL": "mailbox@example.invalid",
}

# Substrings that mark the fatal-error / abort signature we guard against (plus
# any other unexpected interpreter-level failure surfacing at shutdown).
FATAL_MARKERS = (
    "_enter_buffered_busy",
    "Fatal Python error",
    "Abort trap",
    "Traceback (most recent call last)",
)


def _spawn_server() -> subprocess.Popen:
    return subprocess.Popen(
        [sys.executable, str(MAIN_PY)],
        cwd=str(REPO_ROOT),
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env={**os.environ, **SYNTHETIC_ENV},
    )


def _send(proc: subprocess.Popen, message: dict) -> None:
    proc.stdin.write((json.dumps(message) + "\n").encode())
    proc.stdin.flush()


def _read_json_line(proc: subprocess.Popen, timeout: float) -> dict:
    """Read one newline-delimited JSON response, honoring a wall-clock timeout."""
    ready, _, _ = select.select([proc.stdout], [], [], timeout)
    if not ready:
        raise AssertionError("timed out waiting for a server response")
    line = proc.stdout.readline()
    if not line:
        raise AssertionError("server closed stdout before responding")
    return json.loads(line)


def _handshake(proc: subprocess.Popen) -> list[dict]:
    """Run ``initialize`` + ``tools/list``; return the advertised tools.

    Completing this handshake guarantees the anyio stdin worker thread is live
    and blocked in ``readline`` - the precise state that makes finalization race
    the buffered-stream lock.
    """
    _send(
        proc,
        {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2025-06-18",
                "capabilities": {},
                "clientInfo": {"name": "shutdown-probe", "version": "0"},
            },
        },
    )
    init_response = _read_json_line(proc, TIMEOUT_SECONDS)
    assert "result" in init_response, init_response

    _send(proc, {"jsonrpc": "2.0", "method": "notifications/initialized"})
    _send(proc, {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
    tools_response = _read_json_line(proc, TIMEOUT_SECONDS)
    return tools_response.get("result", {}).get("tools", [])


def _wait_for_exit(proc: subprocess.Popen) -> tuple[int, str]:
    """Wait for a self-driven exit; kill and fail if the server hangs."""
    try:
        proc.wait(timeout=TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait()
        stderr = proc.stderr.read().decode(errors="replace")
        raise AssertionError(
            f"server did not exit on its own after disconnect; stderr:\n{stderr}"
        )
    # The process has exited, so both pipes are at EOF: reading cannot deadlock.
    stderr = proc.stderr.read().decode(errors="replace")
    return proc.returncode, stderr


class StdioShutdownRegressionTest(unittest.TestCase):
    def _assert_clean(self, returncode: int, stderr: str) -> None:
        # SIGABRT surfaces as a negative return code (-signal.SIGABRT); demand a
        # clean 0 so any abort or nonzero teardown fails the test.
        self.assertNotEqual(
            returncode, -signal.SIGABRT, f"process aborted (SIGABRT); stderr:\n{stderr}"
        )
        self.assertEqual(
            returncode, 0, f"non-clean exit ({returncode}); stderr:\n{stderr}"
        )
        for marker in FATAL_MARKERS:
            self.assertNotIn(
                marker, stderr, f"unexpected {marker!r} in stderr:\n{stderr}"
            )

    def test_eof_disconnect_exits_cleanly(self) -> None:
        """Normal client disconnect: full handshake, then stdin EOF."""
        for _ in range(CYCLES):
            proc = _spawn_server()
            with proc:
                tools = _handshake(proc)
                self.assertTrue(tools, "server advertised no tools")
                self.assertIn("read_inbox", {t["name"] for t in tools})
                proc.stdin.close()  # EOF = the normal client-disconnect path
                returncode, stderr = _wait_for_exit(proc)
            self._assert_clean(returncode, stderr)

    def test_sigterm_while_connected_exits_cleanly(self) -> None:
        """SIGTERM with stdin still open: the worst-case race.

        stdin is never closed, so the anyio worker stays blocked in ``readline``
        holding the buffer lock. On the unpatched server this is exactly the
        finalization race that aborts; the fix must exit 0.
        """
        for _ in range(CYCLES):
            proc = _spawn_server()
            with proc:
                tools = _handshake(proc)
                self.assertTrue(tools, "server advertised no tools")
                proc.send_signal(signal.SIGTERM)
                returncode, stderr = _wait_for_exit(proc)
            self._assert_clean(returncode, stderr)


if __name__ == "__main__":
    unittest.main()
