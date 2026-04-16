from __future__ import annotations

import json

from pydantic_ai import Agent, ModelSettings
from pydantic_ai.models.groq import GroqModel
from pydantic_ai.providers.groq import GroqProvider

from .config import AgentSettings
from .models import ContentPlan, DeckAnalysis, QAReport, StructurePlan


def _groq_model(settings: AgentSettings):
    model_name = settings.groq_model
    if model_name.startswith("groq:"):
        return model_name
    if settings.groq_api_key:
        provider = GroqProvider(api_key=settings.groq_api_key)
        return GroqModel(model_name, provider=provider)
    return f"groq:{model_name}"


def _analysis_json(analysis: DeckAnalysis) -> str:
    compact = {
        "source_path": analysis.source_path,
        "slide_count": analysis.slide_count,
        "detected_language": analysis.detected_language,
        "slides": [
            {
                "slide_index": slide.slide_index,
                "layout_name": slide.layout_name,
                "placeholder_names": slide.placeholder_names,
                "text_preview": slide.text_preview,
            }
            for slide in analysis.slides
        ],
        "extracted_text": analysis.extracted_text[:8000],
    }
    return json.dumps(compact, ensure_ascii=False, indent=2)


def _fallback_structure_plan() -> StructurePlan:
    return StructurePlan(
        rationale="Fallback sans LLM: aucune operation structurelle.", operations=[]
    )


def _fallback_content_plan(analysis: DeckAnalysis, language: str) -> ContentPlan:
    return ContentPlan(
        language=language,
        slides=[],
    )


def plan_structure(
    settings: AgentSettings,
    analysis: DeckAnalysis,
    user_instruction: str,
) -> StructurePlan:
    prompt = f"""
Tu es un planner de structure PowerPoint.
Objectif: proposer UNIQUEMENT les operations structurelles necessaires (delete/duplicate/add/reorder), sans toucher au contenu texte.

Contraintes:
- Structure d'abord.
- Ne jamais inclure d'operation de contenu.
- Utiliser des index de slides 1-based.
- Si aucune operation n'est necessaire, retourner operations=[] explicitement.
- Favoriser des layouts varies.

Instruction utilisateur:
{user_instruction}

Analyse du template:
{_analysis_json(analysis)}
""".strip()

    agent = Agent(
        _groq_model(settings),
        output_type=StructurePlan,
        model_settings=ModelSettings(temperature=0.1),
        instructions="Tu retournes un JSON strict valide pour StructurePlan.",
    )

    try:
        result = agent.run_sync(prompt)
        return result.output
    except Exception:
        return _fallback_structure_plan()


def plan_content(
    settings: AgentSettings,
    analysis: DeckAnalysis,
    user_instruction: str,
    language: str,
) -> ContentPlan:
    prompt = f"""
Tu es un redacteur de contenu pour slides PowerPoint.
Produis un plan de contenu slide par slide.

Contraintes:
- Langue par defaut: {language}.
- Titres courts et impactants.
- Bullets sous forme d'items separes.
- Si un label inline est necessaire, ecris-le avec un format "Label: valeur".
- Ne retourne que les slides a modifier.
- Ne cree pas d'operations structurelles ici.

Instruction utilisateur:
{user_instruction}

Analyse du template:
{_analysis_json(analysis)}
""".strip()

    agent = Agent(
        _groq_model(settings),
        output_type=ContentPlan,
        model_settings=ModelSettings(temperature=0.2),
        instructions="Tu retournes un JSON strict valide pour ContentPlan.",
    )

    try:
        result = agent.run_sync(prompt)
        output = result.output
        if not output.language:
            output.language = language
        return output
    except Exception:
        return _fallback_content_plan(analysis, language=language)


def plan_content_fixes(
    settings: AgentSettings,
    analysis: DeckAnalysis,
    user_instruction: str,
    language: str,
    reports: list[QAReport],
) -> ContentPlan:
    issues_payload = [
        {
            "mode": report.mode,
            "issues": [issue.model_dump() for issue in report.issues],
            "notes": report.notes,
        }
        for report in reports
    ]

    prompt = f"""
Tu corriges un deck PowerPoint apres QA.
Propose uniquement les modifications de contenu necessaires pour corriger les problemes detectes.

Contraintes:
- Conserver la langue: {language}
- Prioriser les severites high puis medium
- Si aucun correctif de contenu n'est utile, retourner slides=[]
- Pas d'operation structurelle ici

Instruction utilisateur initiale:
{user_instruction}

Issues QA:
{json.dumps(issues_payload, ensure_ascii=False, indent=2)}

Analyse template:
{_analysis_json(analysis)}
""".strip()

    agent = Agent(
        _groq_model(settings),
        output_type=ContentPlan,
        model_settings=ModelSettings(temperature=0.1),
        instructions="Tu retournes un JSON strict valide pour ContentPlan.",
    )

    try:
        result = agent.run_sync(prompt)
        output = result.output
        if not output.language:
            output.language = language
        return output
    except Exception:
        return ContentPlan(language=language, slides=[])
