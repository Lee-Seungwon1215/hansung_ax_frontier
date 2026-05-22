from fastapi import APIRouter

from app import legacy


router = APIRouter(tags=["intranet"])

router.add_api_route("/intranet/dashboard", legacy.intranet_dashboard, methods=["GET"])
router.add_api_route("/auth/me", legacy.auth_me, methods=["GET"])
router.add_api_route("/intranet/approvals", legacy.create_intranet_approval, methods=["POST"])
router.add_api_route(
    "/intranet/approvals/{approval_id}/action",
    legacy.action_intranet_approval,
    methods=["POST"],
)
