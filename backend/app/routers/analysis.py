"""Analysis router — run inference, get results, submit review."""

import asyncio

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.database import get_db, AsyncSessionLocal
from app.inference import run_inference
from app.models import Annotation, Patch, Slide, User
from app.schemas import (
    AnalysisResultResponse, AnalysisRunRequest, AnalysisRunResponse,
    BoundingBox, PatchResult, ReviewRequest, ReviewResponse,
)
from app.routers.slides import get_current_user, require_roles

router = APIRouter(prefix="/analysis", tags=["Analysis"])


# ── Background inference task ───────────────────────────────────────
async def _run_analysis_bg(slide_id: int, threshold: float):
    """Runs in background: update status → run math → save patches."""
    async with AsyncSessionLocal() as db:
        result = await db.execute(select(Slide).where(Slide.id == slide_id))
        slide = result.scalars().first()
        if not slide:
            return

        slide.status = "analyzing"
        await db.commit()

        # Run CPU-bound inference in thread pool
        loop = asyncio.get_running_loop()
        data = await loop.run_in_executor(
            None, run_inference, slide.width, slide.height, threshold
        )

        # Delete old patches if re-running
        old = await db.execute(select(Patch).where(Patch.slide_id == slide_id))
        for p in old.scalars().all():
            await db.delete(p)

        # Save new patches
        for pd in data["patches"]:
            db.add(Patch(
                slide_id=slide_id,
                x_index=pd["x_index"], y_index=pd["y_index"],
                confidence_mild=pd["confidence_mild"],
                confidence_moderate=pd["confidence_moderate"],
                confidence_severe=pd["confidence_severe"],
                confidence_normal=pd["confidence_normal"],
                predicted_grade=pd["predicted_grade"],
                bounding_boxes=pd["bounding_boxes"],
            ))

        slide.status = "processed"
        slide.current_grade = data["overall_grade"]
        slide.overall_confidence = data["overall_confidence"]
        await db.commit()


# ── Run analysis ────────────────────────────────────────────────────
@router.post("/run", response_model=AnalysisRunResponse)
async def run_analysis(
    body: AnalysisRunRequest,
    bg: BackgroundTasks,
    user: User = Depends(require_roles(["Consultant Pathologist", "Resident"])),
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(Slide).where(Slide.id == body.slide_id))
    slide = result.scalars().first()
    if not slide:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Slide not found")
    if slide.status == "analyzing":
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "Already analyzing")

    # Enforce ownership check (IDOR prevention)
    if slide.user_id != user.id:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Access denied to run analysis on this slide.")

    bg.add_task(_run_analysis_bg, slide.id, body.confidence_threshold)

    return AnalysisRunResponse(
        slide_id=slide.id, status="analyzing",
        model_version=body.model_version,
        confidence_threshold=body.confidence_threshold,
    )


# ── Get results ─────────────────────────────────────────────────────
@router.get("/{slide_id}/result", response_model=AnalysisResultResponse)
async def get_result(
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
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Access denied to this slide analysis.")

    patches_q = await db.execute(select(Patch).where(Patch.slide_id == slide_id))
    patches = patches_q.scalars().all()

    patch_results = []
    for p in patches:
        boxes = [BoundingBox(**b) for b in (p.bounding_boxes or [])]
        patch_results.append(PatchResult(
            id=p.id, slide_id=p.slide_id,
            x_index=p.x_index, y_index=p.y_index,
            confidence_mild=p.confidence_mild,
            confidence_moderate=p.confidence_moderate,
            confidence_severe=p.confidence_severe,
            confidence_normal=p.confidence_normal,
            predicted_grade=p.predicted_grade,
            bounding_boxes=boxes,
        ))

    return AnalysisResultResponse(
        slide_id=slide.id,
        overall_grade=slide.current_grade,
        overall_confidence=slide.overall_confidence,
        total_patches=len(patch_results),
        patches=patch_results,
    )


# ── Submit review / annotation ──────────────────────────────────────
@router.put("/{slide_id}/review", response_model=ReviewResponse)
async def submit_review(
    slide_id: int,
    body: ReviewRequest,
    user: User = Depends(require_roles(["Consultant Pathologist", "Resident"])),
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(Slide).where(Slide.id == slide_id))
    slide = result.scalars().first()
    if not slide:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Slide not found")

    # Enforce ownership check (IDOR prevention)
    if slide.user_id != user.id:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Access denied to this slide review.")

    annotation = Annotation(
        slide_id=slide.id, user_id=user.id,
        shape_data=body.annotations,
        final_grade=body.final_grade,
        comments=body.comments,
        icd_10_code=body.icd_10_code,
    )
    db.add(annotation)
    slide.current_grade = body.final_grade
    slide.status = "reviewed"
    await db.commit()

    return ReviewResponse(
        status="Review submitted",
        slide_id=slide.id,
        final_grade=body.final_grade,
        icd_10_code=body.icd_10_code,
        signed_by=user.name,
    )

