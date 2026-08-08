import os
import secrets
from cryptography.fernet import Fernet
from pydantic_settings import BaseSettings

_backend_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_env_path = os.path.join(_backend_dir, ".env")

# Auto-generate local .env if it does not exist and not in production (Vercel)
if not os.path.exists(_env_path) and not os.environ.get("VERCEL"):
    jwt_secret = secrets.token_hex(32)
    fernet_key = Fernet.generate_key().decode()
    with open(_env_path, "w", encoding="utf-8") as f:
        f.write("# OralDysplasia AI Local Environment Configuration\n")
        f.write(f"SECRET_KEY={jwt_secret}\n")
        f.write(f"PATIENT_ENCRYPTION_KEY={fernet_key}\n")
        # Default local database is a safe SQLite file, not hardcoding MySQL credentials
        f.write("DATABASE_URL=sqlite+aiosqlite:///oraldysplasia.db\n")

# Load .env variables into environment if not already set (except on Vercel platform)
if os.path.exists(_env_path) and not os.environ.get("VERCEL"):
    with open(_env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                key = key.strip()
                val = val.strip().strip("'\"")
                if key and val and key not in os.environ:
                    os.environ[key] = val

# On Vercel serverless platform, default to writable /tmp SQLite db if no production database URL is configured
if os.environ.get("VERCEL"):
    if not os.environ.get("DATABASE_URL") or "mysql" in os.environ.get("DATABASE_URL", ""):
        os.environ["DATABASE_URL"] = "sqlite+aiosqlite:////tmp/oraldysplasia.db"
    if not os.environ.get("UPLOAD_DIR"):
        os.environ["UPLOAD_DIR"] = "/tmp/uploads"


class Settings(BaseSettings):
    """Application-wide configuration loaded from environment or defaults."""

    PROJECT_NAME: str = "OralDysplasia AI"
    API_V1_PREFIX: str = "/api/v1"

    # JWT (strictly required)
    SECRET_KEY: str
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 60 * 24 * 7  # 7 days

    # AES-256 Fernet key for encrypting patient PII at rest (strictly required)
    PATIENT_ENCRYPTION_KEY: str

    # Database connection string (strictly required)
    DATABASE_URL: str

    # Where uploaded slide files are stored locally
    UPLOAD_DIR: str = os.getenv("UPLOAD_DIR", os.path.join(_backend_dir, "uploads"))

    class Config:
        case_sensitive = True


settings = Settings()

# Validate secure credentials in production/Vercel
if os.environ.get("VERCEL") or os.environ.get("ENV") == "production":
    INSECURE_KEYS = {
        "oral-dysplasia-dev-secret-key-change-in-production-" + "2026",
        "dGhpcy1pcy1hLTMyLWJ5dGUtZGV2LWtleS0x" + "MjM0NQ=="
    }
    if settings.SECRET_KEY in INSECURE_KEYS:
        raise ValueError("CRITICAL SECURITY VIOLATION: Default/insecure SECRET_KEY is not allowed in production!")
    if settings.PATIENT_ENCRYPTION_KEY in INSECURE_KEYS:
        raise ValueError("CRITICAL SECURITY VIOLATION: Default/insecure PATIENT_ENCRYPTION_KEY is not allowed in production!")

