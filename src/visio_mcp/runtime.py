"""ComWorker: a single STA thread that owns every COM object.

pywin32 COM proxies are apartment-threaded: an object created on one thread
must not be called from another. FastMCP runs tool bodies on asyncio worker
threads, so all COM access is funneled through ONE dedicated thread:

- ThreadPoolExecutor(max_workers=1) guarantees a single, stable thread.
- Its initializer calls pythoncom.CoInitialize() (STA — the canonical
  apartment for Office automation).
- Tools do `await worker.run(fn, ...)`; results/exceptions propagate through
  the Future returned by run_in_executor.
- VisioClient methods must return plain Python data (dicts/lists), never raw
  COM proxies, so nothing COM-bound escapes this thread.
- max_workers=1 also serializes access if the agent issues parallel calls.
- No COM event sinks are registered, so no message pump is needed. If
  WithEvents is ever added, this thread must gain a pump.
"""

from __future__ import annotations

import asyncio
import sys
from concurrent.futures import ThreadPoolExecutor
from functools import partial
from typing import Any, Callable, Optional


def _default_com_initializer() -> None:
    if sys.platform == "win32":
        import pythoncom

        pythoncom.CoInitialize()  # STA


def _default_com_finalizer() -> None:
    if sys.platform == "win32":
        import pythoncom

        pythoncom.CoUninitialize()


class ComWorker:
    def __init__(
        self,
        initializer: Optional[Callable[[], None]] = None,
        finalizer: Optional[Callable[[], None]] = None,
    ):
        self._finalizer = finalizer or _default_com_finalizer
        self._executor = ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="visio-com",
            initializer=initializer or _default_com_initializer,
        )
        self._shut_down = False

    async def run(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        """Run fn(*args, **kwargs) on the COM thread and await its result."""
        if self._shut_down:
            raise RuntimeError("ComWorker has been shut down")
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(self._executor, partial(fn, *args, **kwargs))

    def run_sync(self, fn: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
        """Blocking variant for non-async callers (e.g. the smoke test)."""
        if self._shut_down:
            raise RuntimeError("ComWorker has been shut down")
        return self._executor.submit(partial(fn, *args, **kwargs)).result()

    def shutdown(self, release: Optional[Callable[[], None]] = None) -> None:
        """Release COM references (on the COM thread) and stop the thread.

        `release` should drop all COM object references (e.g. client.release)
        before CoUninitialize runs.
        """
        if self._shut_down:
            return
        self._shut_down = True

        def _fini() -> None:
            try:
                if release is not None:
                    release()
            finally:
                self._finalizer()

        try:
            self._executor.submit(_fini).result(timeout=10)
        except Exception:
            pass  # never let shutdown mask the real work's outcome
        self._executor.shutdown(wait=True)
