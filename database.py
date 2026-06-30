import os

from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

def _default_database_url() -> str:
    if os.getenv("DATABASE_URL"):
        return os.environ["DATABASE_URL"]
    # Vercel serverless 只可寫 /tmp，本機用專案目錄 SQLite
    if os.getenv("VERCEL") or os.getenv("VERCEL_ENV"):
        return "sqlite:////tmp/erp_demo.db"
    return "sqlite:///./erp_demo.db"


DATABASE_URL = _default_database_url()

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}

engine = create_engine(DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


def get_db():
    """資料庫 Session 依賴項"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
