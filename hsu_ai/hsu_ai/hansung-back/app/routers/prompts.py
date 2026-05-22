from fastapi import APIRouter

from app import legacy


router = APIRouter(tags=["prompts"])

router.add_api_route("/prompt-templates", legacy.get_prompt_templates, methods=["GET"])
router.add_api_route(
    "/prompt-templates/{report_type}",
    legacy.update_prompt_template,
    methods=["PUT"],
)
router.add_api_route("/ai-prompts", legacy.get_ai_prompts, methods=["GET"])
router.add_api_route("/ai-prompts/{prompt_key}", legacy.update_ai_prompt, methods=["PUT"])
