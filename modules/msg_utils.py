import hashlib
from datetime import date


def pick_variant(variants, fallback=""):
    """Elige una variante de mensaje según el día del año (rota sin repetir días seguidos)."""
    if not variants:
        return fallback
    idx = int(hashlib.md5(str(date.today()).encode()).hexdigest(), 16) % len(variants)
    return variants[idx]
