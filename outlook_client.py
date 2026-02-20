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
    
    Uses Outlook's Restrict() filter for fast server-side filtering, then
    falls back to full scan if the filter returns no results (e.g. Exchange
    addresses that don't match the SMTP filter).
    
    Args:
        sender_email: The email address to filter by (e.g. "boris@example.com")
        output_dir: Directory to save the text files (defaults to STYLE_SAMPLES_DIR)
        max_count: Maximum number of emails to export
    
    Returns:
        Number of emails exported
    """
    output_dir = output_dir or config.STYLE_SAMPLES_DIR
    os.makedirs(output_dir, exist_ok=True)

    print(f"  [1/5] Connecting to Outlook...")
    outlook = get_outlook()
    namespace = get_namespace(outlook)
    print(f"  [1/5] ✅ Connected to Outlook")

    # Search in Inbox (6 = olFolderInbox)
    print(f"  [2/5] Opening Inbox folder...")
    try:
        inbox = namespace.GetDefaultFolder(6)
        all_messages = inbox.Items
        total_inbox = all_messages.Count
    except Exception as e:
        error_msg = str(e)
        print(f"\n  ❌ Could not open Outlook Inbox!")
        print(f"  Error: {error_msg}")
        if ".ost" in error_msg.lower() or "data file" in error_msg.lower() or "נתונים" in error_msg:
            print(f"\n  🔧 This is an Outlook data file (.ost) problem.")
            print(f"  Try these fixes:")
            print(f"    1. Close Outlook completely, then reopen it")
            print(f"    2. Go to Control Panel → Mail → Data Files and repair")
            print(f"    3. Or delete the .ost file (Outlook will recreate it)")
        else:
            print(f"\n  🔧 Make sure Outlook is open and you're signed in.")
        return 0
    print(f"  [2/5] ✅ Inbox has {total_inbox} total emails")

    sender_lower = sender_email.lower().strip()

    # --- Try fast Restrict() filter first ---
    print(f"  [3/5] Filtering emails from '{sender_email}'...")
    try:
        restriction = f"[SenderEmailAddress] = '{sender_email}'"
        filtered = all_messages.Restrict(restriction)
        filtered_count = filtered.Count
        print(f"  [3/5] ✅ Restrict filter found {filtered_count} emails (SMTP match)")
    except Exception as e:
        print(f"  [3/5] ⚠ Restrict filter failed ({e}), will do full scan")
        filtered_count = 0
        filtered = None

    # If filter found results, use them; otherwise fall back to full scan
    if filtered_count > 0:
        messages_to_scan = filtered
        scan_count = filtered_count
        use_filter = True
    else:
        # Fall back — might be Exchange addresses  
        print(f"  [3/5] ℹ No SMTP matches, falling back to scan (handles Exchange addresses)")
        messages_to_scan = all_messages
        scan_count = total_inbox
        use_filter = False

    messages_to_scan.Sort("[ReceivedTime]", True)  # newest first

    exported = 0
    skipped_existing = 0
    skipped_short = 0
    scanned = 0
    errors = 0

    print(f"  [4/5] Scanning {scan_count} emails...")

    for item in messages_to_scan:
        if exported >= max_count:
            print(f"  [4/5] Reached max count ({max_count}), stopping")
            break

        scanned += 1

        # Progress indicator every 10 emails (more frequent for visibility)
        if scanned % 10 == 0:
            pct = int(scanned / scan_count * 100) if scan_count else 0
            print(f"  [4/5] ... {scanned}/{scan_count} ({pct}%) — exported: {exported}")

        try:
            # If we already used Restrict(), we know the sender matches (for SMTP)
            # but still need to verify for Exchange addresses in full-scan mode
            if not use_filter:
                item_sender = ""
                try:
                    if item.SenderEmailType == "EX":
                        exchg = item.Sender.GetExchangeUser()
                        if exchg:
                            item_sender = exchg.PrimarySmtpAddress or ""
                        else:
                            item_sender = item.SenderEmailAddress or ""
                    else:
                        item_sender = item.SenderEmailAddress or ""
                except Exception:
                    item_sender = item.SenderEmailAddress or ""

                if item_sender.lower().strip() != sender_lower:
                    continue

            subject = item.Subject or "no_subject"
            body = item.Body or ""
            received_time = str(item.ReceivedTime)

            # Skip very short emails
            if len(body.strip()) < 20:
                skipped_short += 1
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

            # Resolve sender info for the file header
            try:
                if use_filter:
                    sender_display = item.SenderEmailAddress or sender_email
                else:
                    sender_display = item_sender
            except Exception:
                sender_display = sender_email

            # Format the email content
            content = (
                f"From: {item.SenderName} <{sender_display}>\n"
                f"Subject: {subject}\n"
                f"Received: {received_time}\n"
                f"{'=' * 50}\n\n"
                f"{body}\n"
            )

            with open(filepath, "w", encoding="utf-8", errors="replace") as f:
                f.write(content)

            exported += 1
            print(f"  [4/5] 📄 Saved: {filename}")

        except Exception as e:
            errors += 1
            if errors <= 5:  # Only show first 5 errors
                print(f"  ⚠ Error on email #{scanned}: {e}")
            continue

    # Summary
    print(f"\n  [5/5] ─── Collection Summary ───")
    print(f"  📊 Scanned: {scanned} emails")
    print(f"  ✅ Exported: {exported} emails")
    if skipped_existing > 0:
        print(f"  ℹ Already exported: {skipped_existing}")
    if skipped_short > 0:
        print(f"  ℹ Skipped (too short): {skipped_short}")
    if errors > 0:
        print(f"  ⚠ Errors: {errors}")

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
