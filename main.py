"""
OritDva - Personalized Email Responder
Main orchestrator: ties together style extraction, Outlook reading, and reply generation.

Usage:
    python main.py collect     - Pull sent emails from Outlook as writing samples
    python main.py extract     - Analyze writing samples and create style profile
    python main.py check       - List unread emails and preview them
    python main.py respond     - Process unread emails and generate draft replies
    python main.py folders     - List available Outlook folders
    python main.py test        - Quick connectivity and API test
"""
import sys
import os

import config
from style_extractor import extract_style, load_style_profile
from outlook_client import get_unread_emails, create_draft_reply, list_folders, export_emails_from_sender
from response_generator import generate_reply_interactive


def cmd_collect():
    """Pull emails from a specific sender in Outlook and save as writing samples."""
    samples_dir = config.STYLE_SAMPLES_DIR

    print("📤 Collecting emails from your Outlook Inbox...\n")

    sender = input("   Enter the sender email address to collect from: ").strip()
    if not sender or "@" not in sender:
        print("  ❌ Invalid email address.")
        return

    try:
        count_input = input("   How many emails to collect? [100]: ").strip()
        max_count = int(count_input) if count_input else 100
    except ValueError:
        max_count = 100

    print(f"\n   Searching Inbox for emails from '{sender}'...")
    print(f"   Saving to: {os.path.abspath(samples_dir)}")

    exported = export_emails_from_sender(
        sender_email=sender,
        output_dir=samples_dir,
        max_count=max_count,
    )

    print(f"\n✅ Exported {exported} emails as writing samples!")
    if exported > 0:
        print(f"\n   Next step: run 'python main.py extract' to build your style profile.")
    else:
        print(f"   No emails found from '{sender}'. Check the address and try again.")


def cmd_extract():
    """Extract style profile from writing samples."""
    samples_dir = config.STYLE_SAMPLES_DIR

    if not os.path.exists(samples_dir):
        os.makedirs(samples_dir, exist_ok=True)
        print(f"📁 Created samples directory: {os.path.abspath(samples_dir)}")
        print(f"   Please add your writing samples (.txt, .eml, .md, .html) to this folder")
        print(f"   Then run: python main.py extract")
        return

    print("🔍 Extracting your writing style...\n")
    profile = extract_style()

    print("\n✅ Style profile created successfully!")
    print(f"   Tone: {profile.get('tone', 'N/A')}")
    print(f"   Formality: {profile.get('formality_level', 'N/A')}/10")

    phrases = profile.get('unique_phrases', [])
    if phrases:
        print(f"   Unique phrases found: {', '.join(phrases[:5])}")


def cmd_check():
    """List unread emails without generating replies."""
    print(f"📬 Checking unread emails in '{config.OUTLOOK_FOLDER}'...\n")

    emails = get_unread_emails(max_count=10)

    if not emails:
        print("  No unread emails found. 🎉")
        return

    print(f"  Found {len(emails)} unread email(s):\n")
    for i, email in enumerate(emails, 1):
        print(f"  {i}. From: {email['sender_name']} <{email['sender_email']}>")
        print(f"     Subject: {email['subject']}")
        print(f"     Received: {email['received_time']}")
        preview = email['body'][:80].replace('\n', ' ').replace('\r', '')
        print(f"     Preview: {preview}...")
        print()


def cmd_respond():
    """Process unread emails: generate replies and save as drafts."""
    # Ensure style profile exists
    try:
        style_profile = load_style_profile()
    except FileNotFoundError:
        print("❌ No style profile found!")
        print("   Run 'python main.py extract' first to analyze your writing samples.")
        return

    print(f"📬 Fetching unread emails from '{config.OUTLOOK_FOLDER}'...\n")
    emails = get_unread_emails(max_count=10)

    if not emails:
        print("  No unread emails found. 🎉")
        return

    print(f"  Found {len(emails)} unread email(s) to process.\n")

    for i, email in enumerate(emails, 1):
        print(f"\n{'='*60}")
        print(f"  [{i}/{len(emails)}]")

        reply_text = generate_reply_interactive(
            email_subject=email['subject'],
            email_body=email['body'],
            sender_name=email['sender_name'],
            style_profile=style_profile,
        )

        if reply_text:
            success = create_draft_reply(email['entry_id'], reply_text)
            if success:
                print("  ✅ Draft saved to Outlook Drafts folder!")
            else:
                print("  ⚠ Could not save draft. Reply text:")
                print(reply_text)
        else:
            print("  ⏭ Skipped.")

    print(f"\n{'='*60}")
    print("✅ Done! Check your Outlook Drafts folder for review.")


def cmd_folders():
    """List available Outlook mail folders."""
    print("📁 Available Outlook folders:\n")
    for folder in list_folders():
        print(f"  • {folder}")


def cmd_test():
    """Quick test of Outlook connection and Gemini API."""
    print("🧪 Running connectivity tests...\n")

    # Test Gemini
    print("1️⃣  Testing Gemini API...")
    if not config.GEMINI_API_KEY:
        print("  ❌ GEMINI_API_KEY not set in .env")
    else:
        try:
            from google import genai
            client = genai.Client(api_key=config.GEMINI_API_KEY)
            response = client.models.generate_content(
                model=config.GEMINI_MODEL,
                contents="Say 'Hello' in one word."
            )
            print(f"  ✅ Gemini API working! Response: {response.text.strip()}")
        except Exception as e:
            print(f"  ❌ Gemini API error: {e}")

    # Test Outlook
    print("\n2️⃣  Testing Outlook connection...")
    try:
        folders = list_folders()
        print(f"  ✅ Outlook connected! Found {len(folders)} folders")
    except Exception as e:
        print(f"  ❌ Outlook error: {e}")

    # Test style profile
    print("\n3️⃣  Checking style profile...")
    try:
        profile = load_style_profile()
        print(f"  ✅ Style profile loaded (tone: {profile.get('tone', 'N/A')})")
    except FileNotFoundError:
        print(f"  ⚠ No style profile yet. Run 'python main.py extract' first.")

    print("\n🧪 Tests complete!")


COMMANDS = {
    "collect": cmd_collect,
    "extract": cmd_extract,
    "check": cmd_check,
    "respond": cmd_respond,
    "folders": cmd_folders,
    "test": cmd_test,
}


def main():
    print()
    print("╔══════════════════════════════════════╗")
    print("║   OritDva - Email Style Responder    ║")
    print("╚══════════════════════════════════════╝")
    print()

    if len(sys.argv) < 2:
        print("Usage: python main.py <command>\n")
        print("Commands:")
        print("  collect  - Pull sent emails from Outlook as writing samples")
        print("  extract  - Analyze writing samples & create style profile")
        print("  check    - List unread emails")
        print("  respond  - Generate draft replies for unread emails")
        print("  folders  - List available Outlook folders")
        print("  test     - Test Outlook & Gemini connectivity")
        return

    command = sys.argv[1].lower()
    if command not in COMMANDS:
        print(f"❌ Unknown command: '{command}'")
        print(f"   Available: {', '.join(COMMANDS.keys())}")
        return

    COMMANDS[command]()


if __name__ == "__main__":
    main()
