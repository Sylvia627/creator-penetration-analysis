# Creator Penetration Analysis Tool

An automated analysis tool for TikTok Shop beauty brand campaigns, powered by Google Gemini AI.

## Background
Built to automate a manual creator penetration analysis workflow at TikTok Shop. The tool replicates and scales an analysis originally done by hand — identifying untapped high-quality beauty creators across L3+, L4+, and curated recommendation lists.

## Features
- Calculates penetration rates across creator tiers (L3+, L4+, SKM Recco)
- Identifies untapped creator pools for outreach
- Generates AI-powered Key Findings using Google Gemini API
- Produces funnel charts (PNG) and a formatted Word report (.docx)

## Setup
```bash
pip install -r requirements.txt
export GEMINI_API_KEY="your-key-here"
python creator_penetration_analysis.py
```

## Input
CSV files with columns: `creator_id`, `is_sampled`, `is_posted`, `gmv`

## Output
- Funnel charts per creator tier
- `Creator_Penetration_Analysis_Q1_2026.docx` with AI-generated insights
