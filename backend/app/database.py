import sys
import os
import asyncio
import pymysql
from sqlalchemy.ext.asyncio import AsyncSession, create_async_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker
from sqlalchemy.engine import make_url

if sys.platform == 'win32':
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except Exception:
        pass

from app.config import settings

# Global variable to track the active database state
DB_STATUS = "unknown"

# ── Dynamic Engine Initialization with XAMPP MySQL Check ──────────────────
def _ensure_db_exists():
    global DB_STATUS
    db_url = settings.DATABASE_URL

    # If DATABASE_URL is already SQLite (set by Vercel api/index.py or env), use it directly
    if db_url.startswith("sqlite"):
        connect_args = {"check_same_thread": False}
        DB_STATUS = "sqlite"
        # Ensure the directory for the SQLite file exists
        try:
            if "///" in db_url:
                db_file = db_url.split("///", 1)[1]
                os.makedirs(os.path.dirname(db_file), exist_ok=True)
        except Exception:
            pass
        return db_url, connect_args

    # If DATABASE_URL is PostgreSQL (production Vercel with Neon/Supabase)
    if db_url.startswith("postgresql") or db_url.startswith("postgres"):
        DB_STATUS = "postgresql"
        return db_url.replace("postgresql://", "postgresql+asyncpg://", 1) if "+asyncpg" not in db_url else db_url, {}

    # MySQL path — try to connect via pymysql (local XAMPP only)
    if db_url.startswith("mysql"):
        try:
            url = make_url(db_url)
            conn = pymysql.connect(
                host=url.host or "127.0.0.1",
                port=url.port or 3306,
                user=url.username or "root",
                password=url.password or "",
                connect_timeout=2
            )
            cursor = conn.cursor()
            db_name = url.database or "oraldysplasia"
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}`")
            conn.commit()
            cursor.close()
            conn.close()
            print(f"[OK] Verified or created XAMPP MySQL database '{db_name}'")
            DB_STATUS = "mysql"
            return db_url, {}
        except Exception as e:
            print(f"[WARN] XAMPP MySQL is not available: {e}")
            print("[INFO] Falling back to local SQLite database...")
            DB_STATUS = "sqlite_fallback"
            backend_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            sqlite_db_path = os.path.join(backend_dir, "oraldysplasia.db").replace("\\", "/")
            return f"sqlite+aiosqlite:///{sqlite_db_path}", {"check_same_thread": False}

    # Unknown — pass through
    DB_STATUS = "other"
    return db_url, {}

# Retrieve connection settings synchronously
_url, _connect_args = _ensure_db_exists()

from sqlalchemy import event
from sqlalchemy.engine import Engine

engine = create_async_engine(
    _url,
    connect_args=_connect_args,
    echo=False,
    pool_pre_ping=False,
    pool_recycle=300,
)

@event.listens_for(Engine, "connect")
def _set_sqlite_pragma(dbapi_connection, connection_record):
    if "sqlite" in _url:
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA synchronous=NORMAL")
        cursor.close()

# ── Session factory ─────────────────────────────────────────────────
AsyncSessionLocal = sessionmaker(
    bind=engine,
    class_=AsyncSession,
    expire_on_commit=False,
)


# ── Base class for all ORM models ───────────────────────────────────
class Base(DeclarativeBase):
    pass


# ── Dependency injection helper ─────────────────────────────────────
async def get_db():
    """Yield an async DB session and handle commit/rollback."""
    async with AsyncSessionLocal() as session:
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
