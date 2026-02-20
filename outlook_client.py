"""
Outlook Client - Reads emails from Outlook and creates draft replies.

Uses win32com to interact with the locally installed Microsoft Outlook application.
"""
import os
import re
import pythoncom
import win32com.client

import config


def get_outlook():
    """Get a reference to the running Outlook application."""
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        raise ConnectionError(
            f"Could not connect to Outlook. Make sure Outlook is running.\n"
            f"Error: {e}"
        )


def get_namespace(outlook):
    """Get the MAPI namespace for mailbox access."""
    return outlook.GetNamespace("MAPI")


def get_unread_emails(folder_name: str = None, max_count: int = 10) -> list[dict]:
    """
    Fetch unread emails from the specified Outlook folder.
    
    Returns a list of dicts with keys:
      - entry_id: unique Outlook identifier for the email
      - subject: email subject line
      - sender_name: display name of the sender
      - sender_email: email address of the sender
      - received_time: when the email was received
      - body: plain text body of the email
      - conversation_id: for threading
    """
    folder_name = folder_name or config.OUTLOOK_FOLDER
    outlook = get_outlook()
    namespace = get_namespace(outlook)

    # Get the default Inbox folder (6 = olFolderInbox)
    if folder_name.lower() == "inbox":
        folder = namespace.GetDefaultFolder(6)
    else:
        # Try to find the folder by name under Inbox
        inbox = namespace.GetDefaultFolder(6)
        try:
            folder = inbox.Folders[folder_name]
        except Exception:
            raise ValueError(
                f"Folder '{folder_name}' not found. Available folders: "
                f"{[f.Name for f in inbox.Folders]}"
            )

    # Filter unread messages
    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # newest first

    unread_filter = "[Unread] = True"
    filtered = messages.Restrict(unread_filter)

    emails = []
    count = 0
    for item in filtered:
        if count >= max_count:
            break
        try:
            # Try to get the actual email address
            sender_email = ""
            try:
                if item.SenderEmailType == "EX":
                    sender_email = item.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    sender_email = item.SenderEmailAddress
            except Exception:
                sender_email = item.SenderEmailAddress or ""

            emails.append({
                "entry_id": item.EntryID,
                "subject": item.Subject or "(No Subject)",
                "sender_name": item.SenderName or "Unknown",
                "sender_email": sender_email,
                "received_time": str(item.ReceivedTime),
                "body": item.Body or "",
                "conversation_id": getattr(item, "ConversationID", ""),
            })
            count += 1
        except Exception as e:
            print(f"  ⚠ Could not read message: {e}")
            continue

    return emails


def get_recent_emails(folder_name: str = None, max_count: int = 10) -> list[dict]:
    """
    Fetch the most recent emails (read or unread) from the specified folder.
    Useful for gathering sent emails as style samples.
    """
    folder_name = folder_name or config.OUTLOOK_FOLDER
    outlook = get_outlook()
    namespace = get_namespace(outlook)

    if folder_name.lower() == "inbox":
        folder = namespace.GetDefaultFolder(6)
    elif folder_name.lower() == "sent" or folder_name.lower() == "sent items":
        folder = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
    else:
        inbox = namespace.GetDefaultFolder(6)
        try:
            folder = inbox.Folders[folder_name]
        except Exception:
            raise ValueError(f"Folder '{folder_name}' not found.")

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    emails = []
    count = 0
    for item in messages:
        if count >= max_count:
            break
        try:
            emails.append({
                "entry_id": item.EntryID,
                "subject": item.Subject or "(No Subject)",
                "sender_name": item.SenderName or "Unknown",
                "sender_email": getattr(item, "SenderEmailAddress", ""),
                "received_time": str(item.ReceivedTime),
                "body": item.Body or "",
            })
            count += 1
        except Exception as e:
            print(f"  ⚠ Could not read message: {e}")
            continue

    return emails


def create_draft_reply(entry_id: str, reply_body: str) -> bool:
    """
    Create a draft reply to the email identified by entry_id.
    The reply is saved as a Draft — NOT sent automatically.
    
    Returns True on success, False on failure.
    """
    outlook = get_outlook()
    namespace = get_namespace(outlook)

    try:
        # Get the original email by its EntryID
        original = namespace.GetItemFromID(entry_id)

        # Create a reply
        reply = original.Reply()
        
        # Set the body — preserve the original conversation below
        reply.Body = reply_body + "\n\n" + reply.Body
        
        # Save as draft (do NOT send)
        reply.Save()

        print(f"  ✓ Draft reply created for: {original.Subject}")
        return True

    except Exception as e:
        print(f"  ✗ Failed to create draft reply: {e}")
        return False


def export_emails_from_sender(
    sender_email: str,
    output_dir: str = None,
    max_count: int = 100,
) -> int:
    """
    Export emails received FROM a specific sender to text files for style analysis.
    
    Searches the Inbox (and optionally other folders) for emails from the 
    given email address and saves them as .txt files.
    
    Args:
        sender_email: The email address to filter by (e.g. "boris@example.com")
        output_dir: Directory to save the text files (defaults to STYLE_SAMPLES_DIR)
        max_count: Maximum number of emails to export
    
    Returns:
        Number of emails exported
    """
    output_dir = output_dir or config.STYLE_SAMPLES_DIR
    os.makedirs(output_dir, exist_ok=True)

    outlook = get_outlook()
    namespace = get_namespace(outlook)

    # Search in Inbox (6 = olFolderInbox)
    inbox = namespace.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)  # newest first

    exported = 0
    skipped_existing = 0
    scanned = 0

    sender_lower = sender_email.lower().strip()

    for item in messages:
        if exported >= max_count:
            break

        scanned += 1
        # Progress indicator every 50 emails
        if scanned % 50 == 0:
            print(f"  ... scanned {scanned} emails, exported {exported} so far")

        try:
            # Resolve the sender's actual SMTP address
            item_sender = ""
            try:
                if item.SenderEmailType == "EX":
                    item_sender = item.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    item_sender = item.SenderEmailAddress
            except Exception:
                item_sender = item.SenderEmailAddress or ""

            # Check if this email is from the target sender
            if item_sender.lower().strip() != sender_lower:
                continue

            subject = item.Subject or "no_subject"
            body = item.Body or ""
            received_time = str(item.ReceivedTime)

            # Skip very short emails
            if len(body.strip()) < 20:
                continue

            # Create a safe filename
            safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:60].strip()
            safe_time = received_time[:10].replace("-", "")  # YYYYMMDD
            filename = f"from_{safe_time}_{safe_subject}.txt"
            filepath = os.path.join(output_dir, filename)

            # Skip if already exported
            if os.path.exists(filepath):
                skipped_existing += 1
                continue

            # Format the email content
            content = (
                f"From: {item.SenderName} <{item_sender}>\n"
                f"Subject: {subject}\n"
                f"Received: {received_time}\n"
                f"{'=' * 50}\n\n"
                f"{body}\n"
            )

            with open(filepath, "w", encoding="utf-8", errors="replace") as f:
                f.write(content)

            exported += 1

        except Exception as e:
            print(f"  ⚠ Could not process message: {e}")
            continue

    print(f"  📊 Scanned {scanned} emails total")
    if skipped_existing > 0:
        print(f"  ℹ Skipped {skipped_existing} already-exported emails")

    return exported


def list_folders() -> list[str]:
    """List all available mail folders for the user."""
    outlook = get_outlook()
    namespace = get_namespace(outlook)
    inbox = namespace.GetDefaultFolder(6)

    folders = ["Inbox", "Sent Items"]
    for folder in inbox.Folders:
        folders.append(folder.Name)
    return folders


if __name__ == "__main__":
    print("📬 Outlook Connection Test")
    print("=" * 40)

    try:
        folders = list_folders()
        print(f"\n📁 Available folders: {folders}")

        print(f"\n📨 Fetching last 3 unread emails from {config.OUTLOOK_FOLDER}...")
        emails = get_unread_emails(max_count=3)

        if not emails:
            print("  No unread emails found.")
        else:
            for email in emails:
                print(f"\n  From: {email['sender_name']} <{email['sender_email']}>")
                print(f"  Subject: {email['subject']}")
                print(f"  Received: {email['received_time']}")
                preview = email['body'][:100].replace('\n', ' ')
                print(f"  Preview: {preview}...")

    except Exception as e:
        print(f"\n❌ Error: {e}")
