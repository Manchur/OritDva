"""
Response Generator - Generates email replies in the user's personal style.

Uses the Style Profile + Gemini to craft replies that sound like the user wrote them.
"""
import json

from google import genai
from google.genai import types

import config
from style_extractor import load_style_profile


REPLY_SYSTEM_PROMPT = """You are a ghostwriter. Your ONLY job is to write email replies 
that perfectly mimic a specific person's writing style. You must sound EXACTLY like them —
not like an AI, not like a generic professional, but like THIS specific person.

Here is their detailed Style Profile:
{style_profile}

CRITICAL RULES:
1. Match their tone EXACTLY — if they're blunt, be blunt. If they're warm, be warm.
2. Use their greeting and closing patterns naturally.
3. Incorporate their unique phrases and terminology where appropriate.
4. Match their punctuation habits (exclamation marks, ellipses, dashes, etc.)
5. Match their paragraph style and sentence structure.
6. Match their formality level ({formality}/10).
7. If they use humor, use similar humor. If they don't, stay serious.
8. NEVER add corporate-speak or AI-style filler unless that's how they write.
9. Keep the reply length similar to how they typically respond.
10. Write in the same language as the incoming email. If the person writes in Hebrew, 
    reply in Hebrew. If English, reply in English.

You will receive an email to reply to. Write ONLY the reply body. 
No subject line. No "Subject:" prefix. Just the reply text.
"""


def generate_reply(
    email_subject: str,
    email_body: str,
    sender_name: str,
    additional_context: str = "",
    style_profile: dict = None,
) -> str:
    """
    Generate a reply to the given email using the user's style profile.
    
    Args:
        email_subject: Subject of the email to reply to
        email_body: Body of the email to reply to
        sender_name: Name of the person who sent the email
        additional_context: Optional context/instructions for this specific reply
        style_profile: Style profile dict (loaded from disk if not provided)
    
    Returns:
        The generated reply text
    """
    if not config.GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY not set. Please add it to your .env file.")

    # Load style profile
    if style_profile is None:
        style_profile = load_style_profile()

    formality = style_profile.get("formality_level", 5)
    profile_json = json.dumps(style_profile, indent=2, ensure_ascii=False)

    # Build system prompt
    system_prompt = REPLY_SYSTEM_PROMPT.format(
        style_profile=profile_json,
        formality=formality,
    )

    # Build user prompt
    user_prompt = f"""Reply to this email:

FROM: {sender_name}
SUBJECT: {email_subject}

--- EMAIL BODY ---
{email_body}
--- END ---
"""
    if additional_context:
        user_prompt += f"\nADDITIONAL CONTEXT/INSTRUCTIONS: {additional_context}\n"

    # Call Gemini
    client = genai.Client(api_key=config.GEMINI_API_KEY)

    response = client.models.generate_content(
        model=config.GEMINI_MODEL,
        contents=user_prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            temperature=0.7,  # some creativity for natural variation
            max_output_tokens=2048,
        )
    )

    return response.text.strip()


def generate_reply_interactive(
    email_subject: str,
    email_body: str,
    sender_name: str,
    style_profile: dict = None,
) -> str:
    """
    Interactive version: shows the email, asks for optional context,
    generates the reply, and lets the user approve or retry.
    """
    if style_profile is None:
        style_profile = load_style_profile()

    print("\n" + "=" * 60)
    print(f"📧 Email from: {sender_name}")
    print(f"   Subject: {email_subject}")
    print("-" * 60)
    # Show first 500 chars of body
    preview = email_body[:500]
    if len(email_body) > 500:
        preview += "..."
    print(preview)
    print("=" * 60)

    context = input("\n💡 Any specific instructions for this reply? (Enter to skip): ").strip()

    while True:
        print("\n🤖 Generating reply...")
        reply = generate_reply(
            email_subject=email_subject,
            email_body=email_body,
            sender_name=sender_name,
            additional_context=context,
            style_profile=style_profile,
        )

        print("\n" + "-" * 60)
        print("📝 DRAFT REPLY:")
        print("-" * 60)
        print(reply)
        print("-" * 60)

        choice = input("\n[A]ccept / [R]etry / [E]dit instructions / [S]kip? ").strip().upper()
        if choice == "A":
            return reply
        elif choice == "R":
            continue
        elif choice == "E":
            context = input("New instructions: ").strip()
            continue
        elif choice == "S":
            return None
        else:
            print("Invalid choice. Please enter A, R, E, or S.")


if __name__ == "__main__":
    # Quick test with a mock email
    test_reply = generate_reply(
        email_subject="Project update meeting",
        email_body="Hi, can we schedule a meeting to discuss the project status? I'm available Tuesday or Wednesday afternoon.",
        sender_name="Test User",
    )
    print("\n📝 Generated reply:")
    print(test_reply)
