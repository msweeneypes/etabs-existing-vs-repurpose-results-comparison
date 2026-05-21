"""
Two-level persistent cache for ETABS comparison results.

Level 1: module-level dict — sub-millisecond within a server session.
Level 2: Viktor Storage API (workspace scope) — survives restarts, shared
         across entities. Falls back silently if Storage is unavailable
         (e.g., unit tests, offline dev without a workspace).

Keys are short MD5 slugs so they fit within the 64-char historical limit.
Values are gzip-compressed pickle to reduce Storage payload size.
"""
import gzip
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
        raw = file_obj.getvalue_binary()
        try:
            value = pickle.loads(gzip.decompress(raw))
        except Exception:
            value = pickle.loads(raw)  # backward compat with uncompressed entries
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
        storage.set(key, File.from_data(gzip.compress(pickle.dumps(value), compresslevel=6)), scope='workspace')
    except Exception:
        pass   # in-process cache still works; Storage unavailable is non-fatal
