"""One process-wide budget per provider/model, shared across batches and retries."""
from contextlib import contextmanager
from threading import BoundedSemaphore, Lock
from urllib.parse import urlsplit

MAX_MODEL_CONCURRENCY = 30
_lock = Lock()
_slots = {}


def worker_count(config):
    return max(1, min(MAX_MODEL_CONCURRENCY, int(config.get('max_workers', 30))))


@contextmanager
def model_slot(api_url, model):
    key = (urlsplit(api_url).netloc.casefold(), model)
    with _lock:
        if key not in _slots:
            _slots[key] = BoundedSemaphore(MAX_MODEL_CONCURRENCY)
        semaphore = _slots[key]
    with semaphore:
        yield
