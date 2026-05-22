from fastapi import APIRouter

from app import legacy


router = APIRouter(tags=["documents"])

router.add_api_route(
    "/documents/{document_id}/download",
    legacy.download_document_template,
    methods=["GET"],
)
