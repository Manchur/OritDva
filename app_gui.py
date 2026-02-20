"""
OritDva - Personalized Email Style Responder (Windows GUI)

A tkinter-based desktop application that wraps the CLI functionality
into a user-friendly Windows interface.
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import os
import sys
import json

# Ensure we can find our modules when running as .exe
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))

import config
from style_extractor import extract_style, load_style_profile, load_samples, build_samples_text
from outlook_client import (
    get_unread_emails, create_draft_reply, list_folders,
    export_emails_from_sender,
)
from response_generator import generate_reply


class LogRedirector:
    """Redirects print() output to the GUI log panel."""

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state="disabled")

    def flush(self):
        pass


class OritDvaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OritDva — Email Style Responder")
        self.root.geometry("900x700")
        self.root.minsize(750, 550)
        self.root.configure(bg="#1e1e2e")

        # Try to set icon (won't crash if missing)
        try:
            self.root.iconbitmap("icon.ico")
        except Exception:
            pass

        self.style_profile = None
        self.current_emails = []
        self.current_email_index = 0

        self._build_styles()
        self._build_ui()
        self._redirect_output()
        self._load_env()

    # ── Styling ──────────────────────────────────────────────
    def _build_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Dark theme colors
        bg = "#1e1e2e"
        surface = "#282840"
        accent = "#7c3aed"
        accent_hover = "#6d28d9"
        text_fg = "#e0e0e0"
        muted = "#888"

        style.configure("TFrame", background=bg)
        style.configure("Surface.TFrame", background=surface)
        style.configure("TLabel", background=bg, foreground=text_fg, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=bg, foreground="#fff",
                        font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", background=bg, foreground=muted,
                        font=("Segoe UI", 9))
        style.configure("Surface.TLabel", background=surface, foreground=text_fg,
                        font=("Segoe UI", 10))

        style.configure("Accent.TButton", background=accent, foreground="white",
                        font=("Segoe UI", 10, "bold"), padding=(16, 8))
        style.map("Accent.TButton",
                  background=[("active", accent_hover), ("disabled", "#444")])

        style.configure("TButton", background=surface, foreground=text_fg,
                        font=("Segoe UI", 10), padding=(12, 6))
        style.map("TButton",
                  background=[("active", "#3a3a5c")])

        style.configure("TEntry", fieldbackground=surface, foreground=text_fg,
                        font=("Segoe UI", 10), padding=6)

        style.configure("TNotebook", background=bg)
        style.configure("TNotebook.Tab", background=surface, foreground=text_fg,
                        font=("Segoe UI", 10), padding=(12, 6))
        style.map("TNotebook.Tab",
                  background=[("selected", accent)],
                  foreground=[("selected", "white")])

    # ── UI Layout ────────────────────────────────────────────
    def _build_ui(self):
        # Header
        header = ttk.Frame(self.root)
        header.pack(fill="x", padx=20, pady=(15, 5))
        ttk.Label(header, text="✉  OritDva", style="Header.TLabel").pack(side="left")
        ttk.Label(header, text="Email Style Responder", style="Sub.TLabel").pack(
            side="left", padx=(10, 0), pady=(8, 0))

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, style="Sub.TLabel")
        status_bar.pack(fill="x", padx=20)

        # Tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=15, pady=10)

        # Tab 1: Setup
        self.setup_tab = ttk.Frame(notebook, style="TFrame")
        notebook.add(self.setup_tab, text="  ⚙  Setup  ")
        self._build_setup_tab()

        # Tab 2: Collect & Extract
        self.collect_tab = ttk.Frame(notebook, style="TFrame")
        notebook.add(self.collect_tab, text="  📤  Collect  ")
        self._build_collect_tab()

        # Tab 3: Respond
        self.respond_tab = ttk.Frame(notebook, style="TFrame")
        notebook.add(self.respond_tab, text="  📧  Respond  ")
        self._build_respond_tab()

        # Tab 4: Log
        self.log_tab = ttk.Frame(notebook, style="TFrame")
        notebook.add(self.log_tab, text="  📋  Log  ")
        self._build_log_tab()

    def _enable_clipboard(self, widget):
        """Enable Ctrl+V, Ctrl+C, Ctrl+A on a widget."""
        def paste(event):
            try:
                text = widget.clipboard_get()
                # If there's a selection, replace it
                try:
                    widget.delete("sel.first", "sel.last")
                except tk.TclError:
                    pass
                widget.insert("insert", text)
            except tk.TclError:
                pass
            return "break"

        def copy(event):
            try:
                widget.clipboard_clear()
                text = widget.selection_get()
                widget.clipboard_append(text)
            except tk.TclError:
                pass
            return "break"

        def select_all(event):
            widget.select_range(0, tk.END)
            widget.icursor(tk.END)
            return "break"

        widget.bind("<Control-v>", paste)
        widget.bind("<Control-V>", paste)
        widget.bind("<Control-c>", copy)
        widget.bind("<Control-C>", copy)
        widget.bind("<Control-a>", select_all)
        widget.bind("<Control-A>", select_all)

    def _build_setup_tab(self):
        f = self.setup_tab
        pad = {"padx": 20, "pady": 5}

        ttk.Label(f, text="Gemini API Key:").pack(anchor="w", **pad)
        self.api_key_var = tk.StringVar(value=config.GEMINI_API_KEY or "")
        key_entry = tk.Entry(
            f, textvariable=self.api_key_var, show="•", width=60,
            bg="#282840", fg="#e0e0e0", insertbackground="#e0e0e0",
            font=("Segoe UI", 10), relief="flat", bd=4,
            selectbackground="#7c3aed", selectforeground="white"
        )
        key_entry.pack(anchor="w", padx=20, pady=(0, 5))
        self._enable_clipboard(key_entry)

        ttk.Label(f, text="Samples Directory:").pack(anchor="w", **pad)
        dir_frame = ttk.Frame(f)
        dir_frame.pack(anchor="w", padx=20, pady=(0, 5))
        self.samples_dir_var = tk.StringVar(value=os.path.abspath(config.STYLE_SAMPLES_DIR))
        ttk.Entry(dir_frame, textvariable=self.samples_dir_var, width=50).pack(side="left")
        ttk.Button(dir_frame, text="Browse...", command=self._browse_samples).pack(
            side="left", padx=(5, 0))

        ttk.Label(f, text="Outlook Folder:").pack(anchor="w", **pad)
        self.folder_var = tk.StringVar(value=config.OUTLOOK_FOLDER)
        ttk.Entry(f, textvariable=self.folder_var, width=30).pack(anchor="w", padx=20, pady=(0, 5))

        btn_frame = ttk.Frame(f)
        btn_frame.pack(anchor="w", padx=20, pady=15)
        ttk.Button(btn_frame, text="💾  Save Settings", style="Accent.TButton",
                   command=self._save_settings).pack(side="left")
        ttk.Button(btn_frame, text="🧪  Test Connection",
                   command=lambda: self._run_async(self._test_connection)).pack(
            side="left", padx=(10, 0))

        # Status profile info
        self.profile_status_var = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.profile_status_var, style="Sub.TLabel").pack(
            anchor="w", padx=20, pady=(10, 0))
        self._check_profile_status()

    def _build_collect_tab(self):
        f = self.collect_tab
        pad = {"padx": 20, "pady": 5}

        ttk.Label(f, text="Collect emails from a specific sender to learn their writing style:"
                  ).pack(anchor="w", **pad)

        ttk.Label(f, text="Sender Email Address:").pack(anchor="w", **pad)
        self.sender_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.sender_var, width=50).pack(
            anchor="w", padx=20, pady=(0, 5))

        ttk.Label(f, text="Max Emails to Collect:").pack(anchor="w", **pad)
        self.max_collect_var = tk.StringVar(value="100")
        ttk.Entry(f, textvariable=self.max_collect_var, width=10).pack(
            anchor="w", padx=20, pady=(0, 10))

        btn_frame = ttk.Frame(f)
        btn_frame.pack(anchor="w", padx=20, pady=5)

        ttk.Button(btn_frame, text="📤  Collect from Outlook", style="Accent.TButton",
                   command=lambda: self._run_async(self._collect_emails)).pack(side="left")
        ttk.Button(btn_frame, text="🔍  Build Style Profile",
                   command=lambda: self._run_async(self._extract_style)).pack(
            side="left", padx=(10, 0))

        # Collect progress
        self.collect_status_var = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.collect_status_var, style="Sub.TLabel").pack(
            anchor="w", padx=20, pady=(10, 0))

    def _build_respond_tab(self):
        f = self.respond_tab

        # Top controls
        ctrl = ttk.Frame(f)
        ctrl.pack(fill="x", padx=20, pady=10)
        ttk.Button(ctrl, text="📬  Fetch Unread Emails", style="Accent.TButton",
                   command=lambda: self._run_async(self._fetch_emails)).pack(side="left")
        self.email_count_var = tk.StringVar(value="No emails loaded")
        ttk.Label(ctrl, textvariable=self.email_count_var, style="Sub.TLabel").pack(
            side="left", padx=(15, 0))

        # Email display
        email_frame = ttk.Frame(f, style="Surface.TFrame")
        email_frame.pack(fill="both", expand=True, padx=20, pady=(0, 5))

        self.email_info_var = tk.StringVar(value="Click 'Fetch Unread Emails' to start")
        ttk.Label(email_frame, textvariable=self.email_info_var, style="Surface.TLabel",
                  wraplength=700).pack(anchor="w", padx=10, pady=5)

        self.email_body_text = scrolledtext.ScrolledText(
            email_frame, height=6, bg="#282840", fg="#e0e0e0",
            font=("Consolas", 9), wrap="word", state="disabled",
            insertbackground="#e0e0e0", relief="flat"
        )
        self.email_body_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Reply section
        reply_frame = ttk.Frame(f)
        reply_frame.pack(fill="x", padx=20, pady=5)

        ttk.Label(reply_frame, text="Additional instructions (optional):").pack(anchor="w")
        self.context_var = tk.StringVar()
        ttk.Entry(reply_frame, textvariable=self.context_var, width=70).pack(
            anchor="w", pady=(0, 5))

        btn_row = ttk.Frame(reply_frame)
        btn_row.pack(anchor="w", pady=5)
        ttk.Button(btn_row, text="🤖  Generate Reply", style="Accent.TButton",
                   command=lambda: self._run_async(self._generate_reply)).pack(side="left")
        ttk.Button(btn_row, text="💾  Save as Draft",
                   command=lambda: self._run_async(self._save_draft)).pack(
            side="left", padx=(10, 0))
        ttk.Button(btn_row, text="⏭  Next Email",
                   command=self._next_email).pack(side="left", padx=(10, 0))

        # Reply display
        self.reply_text = scrolledtext.ScrolledText(
            f, height=6, bg="#282840", fg="#a5f3a5",
            font=("Consolas", 9), wrap="word",
            insertbackground="#e0e0e0", relief="flat"
        )
        self.reply_text.pack(fill="both", expand=True, padx=20, pady=(0, 10))

    def _build_log_tab(self):
        self.log_text = scrolledtext.ScrolledText(
            self.log_tab, bg="#1a1a2e", fg="#aaa",
            font=("Consolas", 9), wrap="word", state="disabled",
            insertbackground="#e0e0e0", relief="flat"
        )
        self.log_text.pack(fill="both", expand=True, padx=15, pady=10)

    # ── Helpers ──────────────────────────────────────────────
    def _redirect_output(self):
        sys.stdout = LogRedirector(self.log_text)
        sys.stderr = LogRedirector(self.log_text)

    def _run_async(self, func):
        """Run a function in a background thread to keep UI responsive."""
        thread = threading.Thread(target=func, daemon=True)
        thread.start()

    def _set_status(self, text):
        self.root.after(0, lambda: self.status_var.set(text))

    def _browse_samples(self):
        path = filedialog.askdirectory(title="Select Samples Directory")
        if path:
            self.samples_dir_var.set(path)

    def _check_profile_status(self):
        path = config.STYLE_PROFILE_PATH
        if os.path.exists(path):
            try:
                profile = load_style_profile(path)
                tone = profile.get("tone", "N/A")
                formality = profile.get("formality_level", "N/A")
                self.profile_status_var.set(
                    f"✅ Style profile loaded — Tone: {tone}, Formality: {formality}/10"
                )
                self.style_profile = profile
            except Exception:
                self.profile_status_var.set("⚠ Style profile exists but could not be loaded")
        else:
            self.profile_status_var.set(
                "⚠ No style profile yet — go to Collect tab to build one"
            )

    def _load_env(self):
        """Load existing .env if present."""
        env_path = os.path.join(os.getcwd(), ".env")
        if os.path.exists(env_path):
            from dotenv import load_dotenv
            load_dotenv(env_path, override=True)
            self.api_key_var.set(os.getenv("GEMINI_API_KEY", ""))
            self.folder_var.set(os.getenv("OUTLOOK_FOLDER", "Inbox"))
            samples = os.getenv("STYLE_SAMPLES_DIR", "./samples")
            self.samples_dir_var.set(os.path.abspath(samples))

    # ── Actions ──────────────────────────────────────────────
    def _save_settings(self):
        """Save settings to .env file."""
        env_path = os.path.join(os.getcwd(), ".env")
        lines = [
            f'GEMINI_API_KEY={self.api_key_var.get().strip()}',
            f'OUTLOOK_FOLDER={self.folder_var.get().strip()}',
            f'STYLE_SAMPLES_DIR={self.samples_dir_var.get().strip()}',
            f'STYLE_PROFILE_PATH=./style_profile.json',
        ]
        with open(env_path, "w") as f:
            f.write("\n".join(lines) + "\n")

        # Update runtime config
        config.GEMINI_API_KEY = self.api_key_var.get().strip()
        config.OUTLOOK_FOLDER = self.folder_var.get().strip()
        config.STYLE_SAMPLES_DIR = self.samples_dir_var.get().strip()

        self._set_status("Settings saved!")
        print("💾 Settings saved to .env")

    def _test_connection(self):
        self._set_status("Testing connections...")
        print("\n🧪 Testing Gemini API...")

        api_key = self.api_key_var.get().strip()
        if not api_key:
            print("  ❌ No API key set")
        else:
            try:
                from google import genai
                client = genai.Client(api_key=api_key)
                response = client.models.generate_content(
                    model=config.GEMINI_MODEL,
                    contents="Say 'Hello' in one word."
                )
                print(f"  ✅ Gemini API working! Response: {response.text.strip()}")
            except Exception as e:
                print(f"  ❌ Gemini error: {e}")

        print("\n🧪 Testing Outlook...")
        try:
            folders = list_folders()
            print(f"  ✅ Outlook connected! Folders: {', '.join(folders)}")
        except Exception as e:
            print(f"  ❌ Outlook error: {e}")

        self._check_profile_status()
        self._set_status("Tests complete")

    def _collect_emails(self):
        sender = self.sender_var.get().strip()
        if not sender or "@" not in sender:
            self._set_status("❌ Enter a valid email address")
            return

        try:
            max_count = int(self.max_collect_var.get())
        except ValueError:
            max_count = 100

        samples_dir = self.samples_dir_var.get().strip()
        self._set_status(f"Collecting emails from {sender}...")
        self.root.after(0, lambda: self.collect_status_var.set("Scanning Inbox..."))

        print(f"\n📤 Searching Inbox for emails from '{sender}'...")
        exported = export_emails_from_sender(
            sender_email=sender,
            output_dir=samples_dir,
            max_count=max_count,
        )

        msg = f"✅ Exported {exported} emails"
        self._set_status(msg)
        self.root.after(0, lambda: self.collect_status_var.set(
            f"{msg}. Click 'Build Style Profile' to analyze them."
        ))
        print(f"\n{msg}")

    def _extract_style(self):
        self._set_status("Building style profile...")
        self.root.after(0, lambda: self.collect_status_var.set("Analyzing writing style with Gemini..."))

        try:
            # Update config with current GUI values
            config.GEMINI_API_KEY = self.api_key_var.get().strip()
            config.STYLE_SAMPLES_DIR = self.samples_dir_var.get().strip()

            profile = extract_style(
                samples_dir=config.STYLE_SAMPLES_DIR,
                output_path=config.STYLE_PROFILE_PATH,
            )
            self.style_profile = profile

            tone = profile.get("tone", "N/A")
            formality = profile.get("formality_level", "N/A")
            msg = f"✅ Style profile ready — Tone: {tone}, Formality: {formality}/10"
            self._set_status(msg)
            self.root.after(0, lambda: self.collect_status_var.set(msg))
            self.root.after(0, self._check_profile_status)

        except Exception as e:
            self._set_status(f"❌ Error: {e}")
            print(f"❌ Style extraction failed: {e}")

    def _fetch_emails(self):
        self._set_status("Fetching unread emails...")
        config.OUTLOOK_FOLDER = self.folder_var.get().strip()

        try:
            self.current_emails = get_unread_emails(
                folder_name=config.OUTLOOK_FOLDER, max_count=20
            )
            self.current_email_index = 0

            count = len(self.current_emails)
            self.root.after(0, lambda: self.email_count_var.set(
                f"{count} unread email(s)" if count else "No unread emails"
            ))

            if count > 0:
                self.root.after(0, self._show_current_email)
                self._set_status(f"Loaded {count} emails")
            else:
                self._set_status("No unread emails found")

        except Exception as e:
            self._set_status(f"❌ Outlook error: {e}")
            print(f"❌ {e}")

    def _show_current_email(self):
        if not self.current_emails:
            return

        email = self.current_emails[self.current_email_index]
        total = len(self.current_emails)
        idx = self.current_email_index + 1

        self.email_info_var.set(
            f"[{idx}/{total}]  From: {email['sender_name']} <{email['sender_email']}>\n"
            f"Subject: {email['subject']}  |  {email['received_time']}"
        )

        self.email_body_text.configure(state="normal")
        self.email_body_text.delete("1.0", tk.END)
        self.email_body_text.insert("1.0", email["body"][:2000])
        self.email_body_text.configure(state="disabled")

        # Clear previous reply
        self.reply_text.delete("1.0", tk.END)

    def _next_email(self):
        if self.current_emails and self.current_email_index < len(self.current_emails) - 1:
            self.current_email_index += 1
            self._show_current_email()

    def _generate_reply(self):
        if not self.current_emails:
            self._set_status("No email selected")
            return

        if not self.style_profile:
            try:
                self.style_profile = load_style_profile()
            except FileNotFoundError:
                self._set_status("❌ No style profile — build one first")
                return

        config.GEMINI_API_KEY = self.api_key_var.get().strip()
        email = self.current_emails[self.current_email_index]
        self._set_status("Generating reply...")

        try:
            reply = generate_reply(
                email_subject=email["subject"],
                email_body=email["body"],
                sender_name=email["sender_name"],
                additional_context=self.context_var.get().strip(),
                style_profile=self.style_profile,
            )

            self.root.after(0, lambda: self._insert_reply(reply))
            self._set_status("Reply generated — review and save as draft")

        except Exception as e:
            self._set_status(f"❌ Generation error: {e}")
            print(f"❌ {e}")

    def _insert_reply(self, text):
        self.reply_text.delete("1.0", tk.END)
        self.reply_text.insert("1.0", text)

    def _save_draft(self):
        if not self.current_emails:
            self._set_status("No email selected")
            return

        reply_body = self.reply_text.get("1.0", tk.END).strip()
        if not reply_body:
            self._set_status("No reply to save")
            return

        email = self.current_emails[self.current_email_index]
        self._set_status("Saving draft...")

        try:
            success = create_draft_reply(email["entry_id"], reply_body)
            if success:
                self._set_status("✅ Draft saved to Outlook!")
                self.root.after(0, self._next_email)
            else:
                self._set_status("❌ Could not save draft")
        except Exception as e:
            self._set_status(f"❌ {e}")


def main():
    root = tk.Tk()
    app = OritDvaApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
