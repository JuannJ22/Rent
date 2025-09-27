from __future__ import annotations

import asyncio
from typing import Callable


class EventBus:
    def __init__(self):
        self.subs: dict[str, list[Callable[[str], None]]] = {}
        self._loop: asyncio.AbstractEventLoop | None = None

    def bind_loop(self, loop: asyncio.AbstractEventLoop | None = None) -> None:
        """Record the main asyncio loop for thread-safe publications."""

        if loop is None:
            try:
                loop = asyncio.get_running_loop()
            except RuntimeError:
                loop = None
        self._loop = loop

    def subscribe(self, topic: str, fn: Callable[[str], None]):
        self.subs.setdefault(topic, []).append(fn)

    def publish(self, topic: str, payload: str):
        loop = self._loop
        if loop is None:
            try:
                loop = asyncio.get_running_loop()
            except RuntimeError:
                loop = None
            else:
                self._loop = loop

        for fn in self.subs.get(topic, []):
            try:
                if loop is None:
                    fn(payload)
                    continue

                try:
                    running = asyncio.get_running_loop()
                except RuntimeError:
                    running = None

                if running is loop:
                    fn(payload)
                else:
                    loop.call_soon_threadsafe(fn, payload)
            except Exception:  # pragma: no cover - defensivo
                pass
