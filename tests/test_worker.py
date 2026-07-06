"""ComWorker: single-thread guarantee, initializer, exception propagation."""

import threading

import pytest

from visio_mcp.runtime import ComWorker


async def test_all_calls_run_on_one_thread_and_initializer_ran_once():
    init_calls = []
    worker = ComWorker(initializer=lambda: init_calls.append(threading.get_ident()),
                       finalizer=lambda: None)
    try:
        idents = set()
        for _ in range(10):
            idents.add(await worker.run(threading.get_ident))
        assert len(idents) == 1, "COM calls leaked onto multiple threads"
        assert init_calls == [next(iter(idents))], "initializer must run once, on the COM thread"
        assert threading.get_ident() not in idents, "COM thread must not be the caller's thread"
    finally:
        worker.shutdown()


async def test_exceptions_propagate_to_awaiter():
    worker = ComWorker(initializer=lambda: None, finalizer=lambda: None)
    try:
        def boom():
            raise RuntimeError("visio exploded")

        with pytest.raises(RuntimeError, match="visio exploded"):
            await worker.run(boom)
        # worker still usable after an exception
        assert await worker.run(lambda: 42) == 42
    finally:
        worker.shutdown()


async def test_shutdown_runs_release_on_com_thread_and_blocks_further_use():
    seen = {}
    worker = ComWorker(initializer=lambda: seen.setdefault("init", threading.get_ident()),
                       finalizer=lambda: seen.setdefault("fini", threading.get_ident()))
    com_thread = await worker.run(threading.get_ident)
    worker.shutdown(release=lambda: seen.setdefault("release", threading.get_ident()))
    assert seen["release"] == com_thread
    assert seen["fini"] == com_thread
    with pytest.raises(RuntimeError):
        await worker.run(lambda: 1)
    worker.shutdown()  # idempotent
