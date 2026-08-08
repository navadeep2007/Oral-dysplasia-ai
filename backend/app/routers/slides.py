"""Slides router — upload, library, detail, dashboard stats."""

import os
from typing import Optional

from fastapi import (
    APIRouter, BackgroundTasks, Depends, File, Form, HTTPException,
    UploadFile, status,
)
from sqlalchemy import func, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.auth import decrypt_pii, encrypt_pii, decode_token
from app.config import settings
from app.database import get_db
from app.models import Slide, User
from app.schemas import DashboardStats, SlideListResponse, SlideResponse

router = APIRouter(prefix="/slides", tags=["Slides"])


# ── Auth dependency ─────────────────────────────────────────────────
from fastapi.security import OAuth2PasswordBearer

_oauth2 = OAuth2PasswordBearer(tokenUrl="api/v1/auth/login")


async def get_current_user(
    token: str = Depends(_oauth2), db: AsyncSession = Depends(get_db)
) -> User:
    payload = decode_token(token)
    if not payload:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid token")
    result = await db.execute(select(User).where(User.email == payload.get("sub")))
    user = result.scalars().first()
    if not user:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "User not found")
    return user


def require_roles(allowed_roles: list[str]):
    async def dependency(user: User = Depends(get_current_user)):
        if user.role not in allowed_roles:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"Insufficient permissions. Required roles: {allowed_roles}"
            )
        return user
    return dependency



# ── Helpers ─────────────────────────────────────────────────────────
def _slide_to_response(s: Slide) -> SlideResponse:
    return SlideResponse(
        id=s.id, user_id=s.user_id,
        patient_id=decrypt_pii(s.patient_id_enc),
        patient_name=decrypt_pii(s.patient_name_enc),
        patient_age=decrypt_pii(s.patient_age_enc) if s.patient_age_enc else None,
        patient_gender=decrypt_pii(s.patient_gender_enc) if s.patient_gender_enc else None,
        anatomical_site=s.anatomical_site,
        filename=s.filename, size_bytes=s.size_bytes,
        width=s.width, height=s.height,
        status=s.status, current_grade=s.current_grade,
        overall_confidence=s.overall_confidence,
        clinical_notes=s.clinical_notes,
        created_at=s.created_at,
    )


# ── Upload ──────────────────────────────────────────────────────────
@router.post("/upload", response_model=SlideResponse)
async def upload_slide(
    file: UploadFile = File(...),
    patient_id: str = Form(...),
    patient_name: str = Form(...),
    patient_age: Optional[str] = Form(None),
    patient_gender: Optional[str] = Form(None),
    anatomical_site: str = Form(...),
    clinical_notes: Optional[str] = Form(None),
    user: User = Depends(require_roles(["Consultant Pathologist", "Resident", "Lab Tech"])),
    db: AsyncSession = Depends(get_db),
):
    allowed = {".svs", ".ndpi", ".tiff", ".tif", ".jpg", ".jpeg", ".png"}
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in allowed:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, f"Unsupported format: {ext}")

    # Save file
    os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
    save_path = os.path.join(settings.UPLOAD_DIR, file.filename)
    content = await file.read()
    with open(save_path, "wb") as f:
        f.write(content)

    slide = Slide(
        user_id=user.id,
        patient_id_enc=encrypt_pii(patient_id),
        patient_name_enc=encrypt_pii(patient_name),
        patient_age_enc=encrypt_pii(patient_age) if patient_age else None,
        patient_gender_enc=encrypt_pii(patient_gender) if patient_gender else None,
        anatomical_site=anatomical_site,
        filename=file.filename,
        file_path=save_path,
        size_bytes=len(content),
        status="uploaded",
        current_grade="pending",
        clinical_notes=clinical_notes,
    )
    db.add(slide)
    await db.commit()
    await db.refresh(slide)
    return _slide_to_response(slide)


# ── Library ─────────────────────────────────────────────────────────
@router.get("/library", response_model=SlideListResponse)
async def slide_library(
    page: int = 1,
    limit: int = 20,
    grade: Optional[str] = None,
    status_filter: Optional[str] = None,
    user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    # Enforce database scoping (IDOR prevention): only retrieve slides uploaded by this user
    q = select(Slide).where(Slide.user_id == user.id).order_by(Slide.created_at.desc())
    count_q = select(func.count(Slide.id)).where(Slide.user_id == user.id)

    if grade:
        q = q.where(Slide.current_grade == grade)
        count_q = count_q.where(Slide.current_grade == grade)
    if status_filter:
        q = q.where(Slide.status == status_filter)
        count_q = count_q.where(Slide.status == status_filter)

    total = (await db.execute(count_q)).scalar() or 0
    q = q.offset((page - 1) * limit).limit(limit)
    rows = (await db.execute(q)).scalars().all()

    return SlideListResponse(
        slides=[_slide_to_response(s) for s in rows],
        total=total, page=page, limit=limit,
    )


# ── Single slide detail ────────────────────────────────────────────
@router.get("/{slide_id}", response_model=SlideResponse)
async def get_slide(
    slide_id: int,
    user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(Slide).where(Slide.id == slide_id))
    slide = result.scalars().first()
    if not slide:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Slide not found")
        
    # Enforce ownership check (IDOR prevention)
    if slide.user_id != user.id:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Access denied to this slide resource")
        
    return _slide_to_response(slide)


# ── Dashboard stats ─────────────────────────────────────────────────
@router.get("/stats/dashboard", response_model=DashboardStats)
async def dashboard_stats(
    user: User = Depends(get_current_user),
    db: AsyncSession = Depends(get_db),
):
    # Scope stats to only the current user's slides
    total = (await db.execute(
        select(func.count(Slide.id)).where(Slide.user_id == user.id)
    )).scalar() or 0
    pending = (await db.execute(
        select(func.count(Slide.id))
        .where(Slide.status.in_(["uploaded", "ready", "preprocessing"]))
        .where(Slide.user_id == user.id)
    )).scalar() or 0
    severe = (await db.execute(
        select(func.count(Slide.id))
        .where(Slide.current_grade == "severe")
        .where(Slide.user_id == user.id)
    )).scalar() or 0
    mild = (await db.execute(
        select(func.count(Slide.id))
        .where(Slide.current_grade == "mild")
        .where(Slide.user_id == user.id)
    )).scalar() or 0
    moderate = (await db.execute(
        select(func.count(Slide.id))
        .where(Slide.current_grade == "moderate")
        .where(Slide.user_id == user.id)
    )).scalar() or 0

    recent_q = select(Slide).where(Slide.user_id == user.id).order_by(Slide.created_at.desc()).limit(5)
    recent = (await db.execute(recent_q)).scalars().all()

    return DashboardStats(
        total_slides=total, pending_review=pending,
        severe_count=severe, mild_count=mild, moderate_count=moderate,
        recent_slides=[_slide_to_response(s) for s in recent],
    )

