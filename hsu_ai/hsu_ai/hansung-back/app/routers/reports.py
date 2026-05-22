from fastapi import APIRouter

from app import legacy


router = APIRouter(tags=["reports"])

router.add_api_route("/generate", legacy.generate_report, methods=["POST"])
router.add_api_route("/save-word", legacy.save_word, methods=["POST"])
router.add_api_route("/save-ppt", legacy.save_ppt, methods=["POST"])
router.add_api_route("/generated-files", legacy.list_generated_files, methods=["GET"])
router.add_api_route(
    "/generated-files/{file_id}/download",
    legacy.download_generated_file,
    methods=["GET"],
)
router.add_api_route("/weekly-reports", legacy.create_weekly_report, methods=["POST"])
router.add_api_route("/weekly-reports", legacy.list_weekly_reports, methods=["GET"])
router.add_api_route(
    "/weekly-reports/compose-request",
    legacy.compose_weekly_report_request,
    methods=["POST"],
)
