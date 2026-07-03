"""Vercel serverless entry — all /api/* requests route here."""
import os
import sys

# Allow imports from project root (main.py, database.py, etc.)
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from main import app  # noqa: E402
