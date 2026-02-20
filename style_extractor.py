"""
Style Extractor - Analyzes writing samples and generates a reusable "Style Profile".

Reads text files from the samples directory, sends them to Gemini for analysis,
and stores the extracted stylistic DNA as a local JSON file.
"""
import json
import os
import glob

from google import genai
from google.genai import types

import config


ANALYSIS_PROMPT = """You are a linguistic analyst. Analyze the following writing samples 
from a single author and produce a comprehensive "Style Profile" in JSON format.

The profile MUST capture ALL of the following dimensions:

1. **tone**: The overall emotional tone (e.g., formal, casual, sarcastic, warm, blunt, diplomatic)
2. **formality_level**: Scale of 1-10 where 1 is extremely casual and 10 is extremely formal
3. **sentence_structure**: How the author builds sentences (short/punchy, long/complex, mixed)
4. **vocabulary_level**: Simple everyday words vs. sophisticated/technical vocabulary
5. **greeting_patterns**: How they open emails/messages (examples from the text)
6. **closing_patterns**: How they sign off (examples from the text)
7. **unique_phrases**: Recurring phrases, pet expressions, or catchphrases
8. **terminology**: Domain-specific or preferred terms they use repeatedly
9. **emotional_expression**: How they express agreement, disagreement, urgency, humor
10. **punctuation_habits**: Use of exclamation marks, ellipses, dashes, parentheses
11. **paragraph_style**: Short paragraphs, long blocks, bullet points, numbered lists
12. **language_quirks**: Any spelling preferences, abbreviations, or unconventional usage
13. **response_patterns**: How they typically structure a reply (acknowledge then answer, jump straight in, etc.)
14. **temper_indicators**: How they handle frustration, pressure, or disagreement in writing
15. **persuasion_style**: How they make arguments or push for action

Return ONLY valid JSON. No markdown fences. No extra text.
Use this exact structure:
{
    "tone": "...",
    "formality_level": 5,
    "sentence_structure": "...",
    "vocabulary_level": "...",
    "greeting_patterns": ["..."],
    "closing_patterns": ["..."],
    "unique_phrases": ["..."],
    "terminology": ["..."],
    "emotional_expression": {
        "agreement": "...",
        "disagreement": "...",
        "urgency": "...",
        "humor": "..."
    },
    "punctuation_habits": "...",
    "paragraph_style": "...",
    "language_quirks": ["..."],
    "response_patterns": "...",
    "temper_indicators": "...",
    "persuasion_style": "...",
    "representative_snippets": ["3-5 short quotes that best represent the author's voice"]
}

--- WRITING SAMPLES ---
{samples}
"""


def load_samples(samples_dir: str) -> list[dict]:
    """Load all text files from the samples directory."""
    samples = []
    patterns = ["*.txt", "*.eml", "*.msg", "*.md", "*.html"]

    for pattern in patterns:
        for filepath in glob.glob(os.path.join(samples_dir, "**", pattern), recursive=True):
            try:
                with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read().strip()
                    if content:
                        samples.append({
                            "filename": os.path.basename(filepath),
                            "content": content
                        })
            except Exception as e:
                print(f"  ⚠ Could not read {filepath}: {e}")

    return samples


def build_samples_text(samples: list[dict]) -> str:
    """Format samples into a single text block for the prompt."""
    parts = []
    for i, sample in enumerate(samples, 1):
        parts.append(f"\n=== SAMPLE {i}: {sample['filename']} ===\n{sample['content']}\n")
    return "\n".join(parts)


def extract_style(samples_dir: str = None, output_path: str = None) -> dict:
    """
    Main function: reads samples, analyzes with Gemini, saves style profile locally.
    Returns the style profile dict.
    """
    samples_dir = samples_dir or config.STYLE_SAMPLES_DIR
    output_path = output_path or config.STYLE_PROFILE_PATH

    if not config.GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY not set. Please add it to your .env file.")

    # Load writing samples
    print(f"📂 Loading samples from: {os.path.abspath(samples_dir)}")
    samples = load_samples(samples_dir)

    if not samples:
        raise FileNotFoundError(
            f"No text files found in '{os.path.abspath(samples_dir)}'. "
            f"Please add .txt, .eml, .msg, .md, or .html files."
        )

    print(f"  ✓ Loaded {len(samples)} writing samples")

    # Build the prompt
    samples_text = build_samples_text(samples)
    prompt = ANALYSIS_PROMPT.format(samples=samples_text)

    # Call Gemini
    print("🤖 Analyzing writing style with Gemini...")
    client = genai.Client(api_key=config.GEMINI_API_KEY)

    response = client.models.generate_content(
        model=config.GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.3,  # low temp for consistent analysis
            max_output_tokens=4096,
        )
    )

    # Parse the JSON response
    raw_text = response.text.strip()
    # Clean potential markdown fences
    if raw_text.startswith("```"):
        raw_text = raw_text.split("\n", 1)[1]
    if raw_text.endswith("```"):
        raw_text = raw_text.rsplit("```", 1)[0]
    raw_text = raw_text.strip()

    try:
        style_profile = json.loads(raw_text)
    except json.JSONDecodeError as e:
        print(f"  ⚠ Failed to parse Gemini response as JSON: {e}")
        print(f"  Raw response:\n{raw_text[:500]}")
        # Save raw text for debugging
        style_profile = {"raw_analysis": raw_text, "parse_error": str(e)}

    # Save locally
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(style_profile, f, indent=2, ensure_ascii=False)

    print(f"  ✓ Style profile saved to: {os.path.abspath(output_path)}")
    return style_profile


def load_style_profile(path: str = None) -> dict:
    """Load a previously saved style profile from disk."""
    path = path or config.STYLE_PROFILE_PATH
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"No style profile found at '{path}'. Run extract_style() first."
        )
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


if __name__ == "__main__":
    profile = extract_style()
    print("\n📋 Style Profile Summary:")
    print(f"  Tone: {profile.get('tone', 'N/A')}")
    print(f"  Formality: {profile.get('formality_level', 'N/A')}/10")
    print(f"  Unique phrases: {profile.get('unique_phrases', [])}")
