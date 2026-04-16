# AI PPTX Editing Agent (Python)

Agent autonome pour modifier des PowerPoint (`.pptx`) a partir d'un template, avec pipeline structure -> contenu -> QA -> export.

## Stack

- Orchestration IA: `pydantic-ai` + provider Groq (`gpt-oss-120b` par defaut)
- Edition PPTX: `python-pptx`
- Operations structurelles PPTX: scripts XML (`unpack`, `clean`, `pack`, `add_slide`)
- QA contenu: `markitdown[pptx]`
- QA visuelle: rendu slides (`soffice` + `pdftoppm`) + Gemini vision (`google.generativeai`)

## Installation

```bash
python -m venv .venv
. .venv/Scripts/activate
pip install -r requirements.txt
```

## Variables d'environnement

Le projet charge automatiquement le fichier `.env` a la racine.

```bash
copy .env.example .env
```

Puis renseigne au minimum `GROQ_API_KEY` dans `.env`.

Contenu type:

```bash
GROQ_API_KEY=...
GROQ_MODEL=gpt-oss-120b

GEMINI_API_KEY=...
GEMINI_VISION_MODEL=gemini-2.0-flash
```

Alternative (variables shell):

```bash
set GROQ_API_KEY=...
set GROQ_MODEL=gpt-oss-120b

set GEMINI_API_KEY=...
set GEMINI_VISION_MODEL=gemini-2.0-flash
```

Optionnel:

```bash
set PPTX_AGENT_WORKDIR=.pptx_agent_work
set PPTX_AGENT_MAX_FIX_LOOPS=2
set PPTX_AGENT_DEFAULT_LANGUAGE=fr
set PPTX_AGENT_ENABLE_VISUAL_QA=true
```

## Lancer le pipeline complet

```bash
python scripts/run_agent.py --input template.pptx --output output.pptx --instruction "Refais le deck pour un pitch produit en francais"
```

Ou avec un fichier d'instructions:

```bash
python scripts/run_agent.py --input template.pptx --output output.pptx --instruction-file brief.md
```

## Scripts utilitaires (inspiration skill Anthropic)

```bash
python scripts/office/unpack.py input.pptx unpacked/
python scripts/add_slide.py unpacked/ slide2.xml
python scripts/clean.py unpacked/
python scripts/office/pack.py unpacked/ output.pptx
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
python scripts/thumbnail.py input.pptx thumbnails.jpg
```

## Notes

- Le pipeline est autonome mais cadre:
  - structure d'abord
  - contenu ensuite
  - QA contenu + QA visuelle
  - boucle de correction avant finalisation
- Sur Colab, `soffice`/`pdftoppm` doivent etre installes pour la QA visuelle.
