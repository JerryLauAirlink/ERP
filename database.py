import os

from sqlalchemy import create_engine, text
from sqlalchemy.orm import declarative_base, sessionmaker

def _default_database_url() -> str:
    if os.getenv("DATABASE_URL"):
        return os.environ["DATABASE_URL"]
    if os.getenv("VERCEL") or os.getenv("VERCEL_ENV"):
        return "sqlite:////tmp/erp_demo.db"
    return "sqlite:///./erp_demo.db"


def _connect_args(url: str) -> dict:
    if url.startswith("sqlite"):
        return {"check_same_thread": False}
    if url.startswith("postgresql") or url.startswith("postgres://"):
        # Supabase requires SSL; connect_timeout avoids hanging on serverless cold start
        return {"sslmode": "require", "connect_timeout": 10}
    return {}


DATABASE_URL = _default_database_url()
connect_args = _connect_args(DATABASE_URL)

engine = create_engine(DATABASE_URL, connect_args=connect_args, pool_pre_ping=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def check_db_connection() -> dict:
    """Ping database; returns status dict for /api/health."""
    try:
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        provider = "supabase" if "supabase" in DATABASE_URL else (
            "postgresql" if DATABASE_URL.startswith("postgresql") else "sqlite"
        )
        return {"ok": True, "provider": provider}
    except Exception as exc:
        return {"ok": False, "provider": "unknown", "error": str(exc)}
