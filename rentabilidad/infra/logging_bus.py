class EventBus:
    def __init__(self):
        self.subs: dict[str, list] = {}

    def subscribe(self, topic: str, fn):
        self.subs.setdefault(topic, []).append(fn)

    def publish(self, topic: str, payload: str):
        for fn in self.subs.get(topic, []):
            try:
                fn(payload)
            except Exception:  # pragma: no cover - defensivo
                pass
