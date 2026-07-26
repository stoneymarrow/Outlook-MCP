"""Regression coverage for stdio shutdown: never abort, never wedge.

The server drives stdin/stdout through an anyio worker thread that blocks in
``readline`` holding the ``BufferedReader`` lock and cannot be cancelled. That
one fact causes two distinct failures, and ``main._run_stdio`` must prevent
both:

* Abort - if interpreter finalization runs while the worker holds the lock,
  CPython's ``_enter_buffered_busy`` fatal-errors and ``abort()``s the process
  (SIGABRT), producing a macOS crash report. Guarded by exiting through
  ``os._exit`` instead of finalizing.
* Wedge - on SIGINT, anyio cancels the task group and then waits forever on
  that uncancellable worker, so ``mcp.run`` never returns and the ``os._exit``
  after it is unreachable. Observed directly: a SIGINT'd server sat in
  ``kevent`` indefinitely while the worker sat in ``read``. Guarded by the
  watchdog that ``_run_stdio`` arms from its signal handler.

These probes are fully contained: the server subprocess runs with synthetic
environment values, speaks only the MCP ``initialize`` / ``tools/list``
handshake (no Graph calls, no real credentials, ``.env`` unread), and is driven
over an in-process pipe. The EOF and SIGTERM disconnects are repeated many
times because that fault is an intermittent race - a single clean cycle proves
nothing. The SIGINT wedge is deterministic, so it needs only a few cycles.
"""

import json
import os
import select
import signal
import subprocess
import sys
import time
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
MAIN_PY = REPO_ROOT / "main.py"

# Enough isolated cycles that an intermittent finalization race would surface.
CYCLES = 15
# The SIGINT wedge reproduces every time, so a few cycles are enough - and each
# one costs the full watchdog grace period.
SIGINT_CYCLES = 3
TIMEOUT_SECONDS = 15.0

# The watchdog grace plus room for process teardown. Comfortably under
# TIMEOUT_SECONDS so a wedge fails this bound rather than the hang timeout,
# which keeps the failure message pointed at the right defect.
SIGINT_EXIT_BUDGET_SECONDS = 8.0

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

    def test_sigint_while_connected_exits_promptly(self) -> None:
        """SIGINT with stdin still open: the wedge, not the abort.

        anyio cancels the task group and then waits on the stdin worker, which
        ``to_thread.run_sync`` cannot cancel while it is blocked in ``read``.
        Without the watchdog ``mcp.run`` never returns, so the ``os._exit``
        after it is unreachable and the process runs forever - it is killed by
        the harness rather than exiting. The watchdog must bound that.
        """
        for _ in range(SIGINT_CYCLES):
            proc = _spawn_server()
            with proc:
                tools = _handshake(proc)
                self.assertTrue(tools, "server advertised no tools")
                started = time.monotonic()
                proc.send_signal(signal.SIGINT)
                returncode, stderr = _wait_for_exit(proc)
                elapsed = time.monotonic() - started
            self._assert_clean(returncode, stderr)
            self.assertLess(
                elapsed,
                SIGINT_EXIT_BUDGET_SECONDS,
                f"SIGINT shutdown took {elapsed:.1f}s - the unwind is wedged on "
                "the uncancellable stdin read and the watchdog did not bound it",
            )


if __name__ == "__main__":
    unittest.main()
