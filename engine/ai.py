"""
Google Gemini (gemini-2.5-flash) integration for auto-generating the
event report description.

Uses the plain REST endpoint via `requests` so no extra SDK dependency
is required. The generated text is a finished, one-page-max report
description that:

  * matches the tone/structure of a reference report (when provided),
  * marks important words with **bold** markers, which the template
    engine converts into real bold runs,
  * returns multiple paragraphs separated by blank lines.

API key resolution order:
  1. `GEMINI_API_KEY` environment variable
  2. `GEMINI_API_KEY` inside `.streamlit/secrets.toml`
  3. an explicit `api_key` argument (e.g. pasted into the UI)
"""

import os
from pathlib import Path

import requests

MODELS = [
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
]

_ENDPOINT = (
    "https://generativelanguage.googleapis.com/v1beta/"
    "models/{model}:generateContent"
)

_MAX_WORDS = 320  # roughly a single page


def resolve_api_key():
    """
    Return the configured Gemini API key, or None.

    Resolution order:
      1. Streamlit secrets (`st.secrets["GEMINI_API_KEY"]`) - works on
         Streamlit Community Cloud and locally via
         .streamlit/secrets.toml.
      2. GEMINI_API_KEY environment variable.
      3. A GEMINI_API_KEY line inside the local .streamlit/secrets.toml
         (used when the CLI runs without Streamlit initialised).
    """
    try:
        import streamlit as st

        key = (st.secrets.get("GEMINI_API_KEY") or "").strip()
        if key:
            return key
    except Exception:
        pass

    env_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if env_key:
        return env_key

    secrets_path = (
        Path(__file__).resolve().parent.parent
        / ".streamlit"
        / "secrets.toml"
    )
    try:
        for line in secrets_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if "=" in line:
                key, _, value = line.partition("=")
                key = key.strip()
                value = value.strip().strip("\"'")
                if key == "GEMINI_API_KEY" and value:
                    return value
    except OSError:
        pass

    return None


def build_description_prompt(inputs):
    """
    Build the prompt used to generate the report description.

    inputs may contain: title, date_time, venue and an optional
    reference_description (style reference from an existing report).
    """
    title = (inputs.get("title") or "").strip()
    date_time = (inputs.get("date_time") or "").strip()
    venue = (inputs.get("venue") or "").strip()
    reference = (inputs.get("reference_description") or "").strip()

    lines = [
        "You are writing the body text of a professional academic "
        "event report for an Institution's Innovation Council (IIC) / "
        "Computer Science department.",
        "",
        "Write the report description in ENGLISH, exactly following this "
        "structure and tone:",
        "1. First paragraph - welcome / context: name the event, its "
        "date and venue, who organized it (student association), and the "
        "topic it covered.",
        "2. Second paragraph - the session itself: what was taught, "
        "the main tools/technologies/concepts demonstrated, the "
        "speaker/resource person, and one example activity.",
        "3. Third paragraph - outcome: skills the participants gained, "
        "and a closing sentence saying the event was successful and "
        "well received.",
        "",
        "Rules:",
        "- Keep the total length under " + str(_MAX_WORDS) + " words "
        "(MUST fit within one printed page).",
        "- Put **double asterisks around the most important words and "
        "phrases** (event name, key technologies, skills, organizer) so "
        "they are bolded in the document.",
        "- Separate the three paragraphs with a blank line.",
        "- Do NOT add a heading, do NOT add bullet lists, do NOT include "
        "any instructions or commentary.",
    ]

    if reference:
        lines += [
            "",
            "Match the tone, vocabulary and level of detail of this "
            "reference report, but do not copy sentences verbatim:",
            "",
        ]
        lines.append("REFERENCE REPORT START")
        lines.append(reference)
        lines.append("REFERENCE REPORT END")

    lines += ["", "EVENT DETAILS", ""]
    lines.append(f"Event title   : {title or 'N/A'}")
    lines.append(f"Date & time   : {date_time or 'N/A'}")
    lines.append(f"Venue         : {venue or 'N/A'}")

    lines += [
        "",
        "Return ONLY the finished description text.",
    ]

    return "\n".join(lines)


def call_gemini(prompt, api_key, model=None):
    """
    Send `prompt` to Gemini and return the generated text.

    Tries each model in MODELS (starting at `model`) until one answers.
    Raises ValueError with a readable message on failure.
    """
    models = (
        [model] + [m for m in MODELS if m != model]
        if model in MODELS
        else MODELS
    )

    last_error = None

    for candidate in models:
        response = requests.post(
            f"{_ENDPOINT.format(model=candidate)}?key={api_key}",
            json={
                "contents": [
                    {
                        "role": "user",
                        "parts": [{"text": prompt}],
                    }
                ],
                "generationConfig": {
                    "temperature": 0.7,
                    "maxOutputTokens": 1200,
                },
            },
            timeout=60,
        )

        if response.status_code == 200:
            data = response.json()
            try:
                text = (
                    data["candidates"][0]
                    ["content"]["parts"][0]["text"]
                )
            except (KeyError, IndexError, TypeError):
                last_error = "Unexpected response from Gemini."
                continue
            return text.strip()

        if response.status_code == 404:
            last_error = f"Model {candidate!r} unavailable."
            continue

        try:
            message = (
                response.json().get("error", {}).get("message")
                or f"HTTP {response.status_code}"
            )
        except ValueError:
            message = f"HTTP {response.status_code}"

        raise ValueError(message)

    raise ValueError(last_error or "Gemini request failed.")


def generate_description(inputs, api_key=None):
    """
    High-level helper: resolve the API key and produce the description.

    Returns the generated text. Raises ValueError when no key is
    configured or the request fails.
    """
    key = api_key or resolve_api_key()

    if not key:
        raise ValueError(
            "No Gemini API key found. Set GEMINI_API_KEY in your "
            "environment, add it to .streamlit/secrets.toml, or paste "
            "it in the app."
        )

    prompt = build_description_prompt(inputs)

    return call_gemini(prompt, key)