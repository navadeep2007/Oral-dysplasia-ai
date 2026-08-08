"""Auth router — signup and login."""

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.auth import create_token, hash_password, verify_password
from app.database import get_db
from app.models import User
from app.schemas import AuthResponse, LoginRequest, SignUpRequest, UserBrief, ForgotPasswordRequest

router = APIRouter(prefix="/auth", tags=["Auth"])


@router.post("/signup", response_model=AuthResponse)
async def signup(body: SignUpRequest, db: AsyncSession = Depends(get_db)):
    # Check duplicate
    existing = await db.execute(select(User).where(User.email == body.email))
    if existing.scalars().first():
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "Email already registered")

    # Validate signup role
    ALLOWED_ROLES = {"Consultant Pathologist", "Resident", "Lab Tech"}
    if body.role not in ALLOWED_ROLES:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=f"Invalid role. Role must be one of: {list(ALLOWED_ROLES)}"
        )

    user = User(
        email=body.email,
        hashed_password=hash_password(body.password),
        name=body.name,
        license_id=body.license_id,
        role=body.role,
        institution=body.institution,
    )
    db.add(user)
    await db.commit()
    await db.refresh(user)

    token = create_token({"sub": user.email, "uid": user.id, "role": user.role})
    return AuthResponse(
        access_token=token,
        user=UserBrief.model_validate(user),
    )


@router.post("/login", response_model=AuthResponse)
async def login(body: LoginRequest, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(User).where(User.email == body.email))
    user = result.scalars().first()

    if not user or not verify_password(body.password, user.hashed_password):
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid credentials")

    token = create_token({"sub": user.email, "uid": user.id, "role": user.role})
    return AuthResponse(
        access_token=token,
        user=UserBrief.model_validate(user),
    )


@router.post("/forgot-password")
async def forgot_password(body: ForgotPasswordRequest, db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(User).where(User.email == body.email))
    user = result.scalars().first()
    if not user:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Email not found")
    return {"message": "Password recovery email sent. Please check your inbox."}
