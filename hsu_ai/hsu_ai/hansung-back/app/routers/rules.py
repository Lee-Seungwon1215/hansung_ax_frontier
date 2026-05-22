from fastapi import APIRouter

from app import legacy


router = APIRouter(tags=["rules"])

router.add_api_route("/rules/search", legacy.search_rule_documents, methods=["GET"])
