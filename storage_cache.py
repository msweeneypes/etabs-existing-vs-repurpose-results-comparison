"""
Two-level persistent cache for ETABS comparison results.

Level 1: module-level dict — sub-millisecond within a server session.
Level 2: Viktor Storage API (workspace scope) — survives restarts, shared
         across entities. Falls back silently if Storage is unavailable
         (e.g., unit tests, offline dev without a workspace).

Keys are short MD5 slugs so they fit within the 64-char historical limit.
"""
import hashlib
import pickle

_MEMORY: dict = {}


def _storage_key(prefix: str, cache_key) -> str:
    h = hashlib.md5(str(cache_key).encode()).hexdigest()
    return f'{prefix}_{h}'


def get_cached(prefix: str, cache_key):
    """Return the cached value or None if not found at either level."""
    mem_key = (prefix, cache_key)
    if mem_key in _MEMORY:
        return _MEMORY[mem_key]

    try:
        from viktor import Storage
        storage = Storage()
        key = _storage_key(prefix, cache_key)
        file_obj = storage.get(key, scope='workspace')
        value = pickle.loads(file_obj.getvalue_binary())
        _MEMORY[mem_key] = value   # warm the in-process cache for this session
        return value
    except Exception:
        return None


def set_cached(prefix: str, cache_key, value) -> None:
    """Write value to both the in-process cache and Viktor Storage."""
    mem_key = (prefix, cache_key)
    _MEMORY[mem_key] = value

    try:
        from viktor import Storage, File
        storage = Storage()
        key = _storage_key(prefix, cache_key)
        storage.set(key, File.from_data(pickle.dumps(value)), scope='workspace')
    except Exception:
        pass   # in-process cache still works; Storage unavailable is non-fatal
