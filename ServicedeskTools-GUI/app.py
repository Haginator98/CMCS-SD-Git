from pathlib import Path

import customtkinter as ctk

from script_registry import (
    SCRIPT_CATEGORIES,
    CATEGORY_DESCRIPTIONS,
    CATEGORY_ICONS,
    get_script_path,
    script_exists,
    find_pwsh,
)
from terminal_widget import TerminalWidget
from module_checker import ModuleChecker
from github_sync import GitHubSyncer
from update_checker import check_for_update, open_releases_page


class ServicedeskApp(ctk.CTk):
    """Main application window for Servicedesk Tools."""

    APP_TITLE = "Servicedesk Tools"
    APP_VERSION = "0.9"

    def __init__(self):
        super().__init__()

        # Window setup
        self.title(f"{self.APP_TITLE} v{self.APP_VERSION}")
        self.geometry("1100x720")
        self.minsize(900, 600)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.pwsh_path = find_pwsh() or "pwsh"
        self.module_checker = ModuleChecker()
        self.syncer = GitHubSyncer()
        self._selected_category: str | None = None
        self._category_buttons: dict[str, ctk.CTkButton] = {}

        self._build_ui()
        self._check_pwsh()

        # Auto-sync scripts from GitHub on startup
        self._sync_scripts()

        # Check for app updates
        check_for_update(self.APP_VERSION, callback=self._on_update_check)

    # ──────────────────────────── UI LAYOUT ────────────────────────────

    def _build_ui(self):
        # Grid layout: sidebar | content
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Sidebar ──
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#181825")
        self.sidebar.grid(row=0, column=0, sticky="nswe")
        self.sidebar.grid_propagate(False)

        # App title in sidebar
        title_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        title_frame.pack(fill="x", padx=16, pady=(20, 4))

        ctk.CTkLabel(
            title_frame,
            text="🛠  Servicedesk Tools",
            font=ctk.CTkFont(size=17, weight="bold"),
            text_color="#cdd6f4",
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_frame,
            text=f"v{self.APP_VERSION} — Mr. Hagen",
            font=ctk.CTkFont(size=11),
            text_color="#6c7086",
        ).pack(anchor="w", pady=(2, 0))

        # Update banner (hidden by default)
        self.update_banner = ctk.CTkButton(
            self.sidebar,
            text="",
            font=ctk.CTkFont(size=11, weight="bold"),
            height=32,
            corner_radius=8,
            fg_color="#fab387",
            hover_color="#e8a070",
            text_color="#1e1e2e",
            command=open_releases_page,
        )
        # Not packed yet — shown only when update is available

        # Separator
        ctk.CTkFrame(self.sidebar, height=1, fg_color="#313244").pack(fill="x", padx=16, pady=(12, 8))

        # Category label
        ctk.CTkLabel(
            self.sidebar,
            text="CATEGORIES",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#6c7086",
        ).pack(anchor="w", padx=20, pady=(8, 4))

        # Category buttons
        for cat_name in SCRIPT_CATEGORIES:
            icon = CATEGORY_ICONS.get(cat_name, "📂")
            desc = CATEGORY_DESCRIPTIONS.get(cat_name, cat_name)
            btn = ctk.CTkButton(
                self.sidebar,
                text=f" {icon}  {cat_name}",
                font=ctk.CTkFont(size=14),
                anchor="w",
                height=38,
                corner_radius=8,
                fg_color="transparent",
                hover_color="#313244",
                text_color="#cdd6f4",
                command=lambda c=cat_name: self._select_category(c),
            )
            btn.pack(fill="x", padx=12, pady=2)
            self._category_buttons[cat_name] = btn

        # Spacer
        ctk.CTkFrame(self.sidebar, fg_color="transparent").pack(fill="both", expand=True)

        # Sync button
        self.sync_btn = ctk.CTkButton(
            self.sidebar,
            text="🔄  Sync Scripts",
            font=ctk.CTkFont(size=12),
            height=34,
            corner_radius=8,
            fg_color="#313244",
            hover_color="#45475a",
            text_color="#a6adc8",
            command=self._sync_scripts,
        )
        self.sync_btn.pack(fill="x", padx=12, pady=(4, 4))

        # Module status button at bottom
        self.module_status_btn = ctk.CTkButton(
            self.sidebar,
            text="⚙  Check Modules",
            font=ctk.CTkFont(size=12),
            height=34,
            corner_radius=8,
            fg_color="#313244",
            hover_color="#45475a",
            text_color="#a6adc8",
            command=self._check_modules,
        )
        self.module_status_btn.pack(fill="x", padx=12, pady=(4, 8))

        # PowerShell status
        self.pwsh_label = ctk.CTkLabel(
            self.sidebar,
            text="",
            font=ctk.CTkFont(size=10),
            text_color="#6c7086",
        )
        self.pwsh_label.pack(padx=16, pady=(0, 16))

        # ── Main content area ──
        self.content = ctk.CTkFrame(self, fg_color="#1e1e2e", corner_radius=0)
        self.content.grid(row=0, column=1, sticky="nswe")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(1, weight=1)

        # Header area (category info + script list)
        self.header_frame = ctk.CTkFrame(self.content, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="nswe", padx=16, pady=(16, 8))
        self.header_frame.grid_columnconfigure(0, weight=1)

        self.header_label = ctk.CTkLabel(
            self.header_frame,
            text="Select a category to get started",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#cdd6f4",
        )
        self.header_label.grid(row=0, column=0, sticky="w")

        self.header_desc = ctk.CTkLabel(
            self.header_frame,
            text="Choose a category from the sidebar to see available scripts.",
            font=ctk.CTkFont(size=13),
            text_color="#6c7086",
        )
        self.header_desc.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # Script list (scrollable)
        self.script_scroll = ctk.CTkScrollableFrame(
            self.content,
            fg_color="transparent",
            corner_radius=0,
        )
        self.script_scroll.grid(row=1, column=0, sticky="nswe", padx=12, pady=(4, 4))
        self.script_scroll.grid_columnconfigure(0, weight=1)

        # Terminal area
        self.terminal_label = ctk.CTkLabel(
            self.content,
            text="OUTPUT",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#6c7086",
        )
        self.terminal_label.grid(row=2, column=0, sticky="w", padx=20, pady=(8, 2))

        self.terminal = TerminalWidget(self.content, fg_color="transparent")
        self.terminal.grid(row=3, column=0, sticky="nswe", padx=12, pady=(0, 12))

        # Give terminal ~40% of vertical space
        self.content.grid_rowconfigure(1, weight=3)
        self.content.grid_rowconfigure(3, weight=4)

        # Welcome message
        self.terminal.append_output(
            "Welcome to Servicedesk Tools!\n"
            "An idea by Servicedesk, made by Mr. Hagen — 2025/2026\n\n"
            "Scripts are synced from GitHub on startup.\n"
            "Remember: You need to have PIM activated (User/Exchange PIM recommended).\n"
        )

    # ──────────────────────────── CATEGORY SELECTION ────────────────────────────

    def _select_category(self, category: str):
        self._selected_category = category

        # Update button highlighting
        for cat, btn in self._category_buttons.items():
            if cat == category:
                btn.configure(fg_color="#45475a", text_color="#cdd6f4")
            else:
                btn.configure(fg_color="transparent", text_color="#cdd6f4")

        icon = CATEGORY_ICONS.get(category, "📂")
        desc = CATEGORY_DESCRIPTIONS.get(category, category)
        self.header_label.configure(text=f"{icon}  {category}")
        self.header_desc.configure(text=desc)

        # Clear and repopulate script list
        for widget in self.script_scroll.winfo_children():
            widget.destroy()

        scripts = SCRIPT_CATEGORIES.get(category, [])
        for i, script_info in enumerate(sorted(scripts, key=lambda s: s["name"])):
            exists = script_exists(script_info["script"])
            script_path = get_script_path(script_info["script"])

            card = ctk.CTkFrame(
                self.script_scroll,
                fg_color="#313244" if exists else "#2a2637",
                corner_radius=10,
                height=48,
            )
            card.grid(row=i, column=0, sticky="ew", pady=3)
            card.grid_columnconfigure(1, weight=1)

            status_icon = "✅" if exists else "❌"
            ctk.CTkLabel(
                card,
                text=status_icon,
                font=ctk.CTkFont(size=14),
                width=30,
            ).grid(row=0, column=0, padx=(12, 4), pady=8)

            ctk.CTkLabel(
                card,
                text=script_info["name"],
                font=ctk.CTkFont(size=14),
                text_color="#cdd6f4" if exists else "#6c7086",
                anchor="w",
            ).grid(row=0, column=1, sticky="w", pady=8)

            if exists:
                run_btn = ctk.CTkButton(
                    card,
                    text="▶ Run",
                    width=70,
                    height=30,
                    corner_radius=6,
                    fg_color="#a6e3a1",
                    hover_color="#80d090",
                    text_color="#1e1e2e",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    command=lambda p=script_path: self._run_script(p),
                )
                run_btn.grid(row=0, column=2, padx=(8, 12), pady=8)
            else:
                ctk.CTkLabel(
                    card,
                    text="Not found",
                    font=ctk.CTkFont(size=11),
                    text_color="#f38ba8",
                ).grid(row=0, column=2, padx=(8, 12), pady=8)

    # ──────────────────────────── ACTIONS ────────────────────────────

    def _run_script(self, script_path: Path):
        if self.terminal.is_running:
            self.terminal.append_output("\n⚠ A script is already running. Stop it first.\n")
            return
        self.terminal.run_script(script_path, self.pwsh_path)

    def _check_pwsh(self):
        pwsh = find_pwsh()
        if pwsh:
            self.pwsh_path = pwsh
            self.pwsh_label.configure(text=f"✅ pwsh found", text_color="#a6e3a1")
        else:
            self.pwsh_label.configure(text="❌ pwsh not found — install PowerShell", text_color="#f38ba8")

    def _check_modules(self):
        self.terminal.clear_output()
        self.terminal.append_output("⚙ Checking PowerShell modules...\n\n")
        self.module_checker.check_modules(callback=self._on_modules_checked)

    def _on_modules_checked(self, modules):
        def _update():
            if not self.module_checker.pwsh_available:
                self.terminal.append_output(
                    "❌ PowerShell (pwsh) not found.\n"
                    "Install with: brew install powershell/tap/powershell\n"
                )
                return

            all_ok = True
            for mod in modules:
                if mod.installed:
                    self.terminal.append_output(f"  ✅ {mod.name} ({mod.version})\n")
                else:
                    self.terminal.append_output(f"  ❌ {mod.name} — not installed\n")
                    all_ok = False

            self.terminal.append_output(f"\n{'─' * 50}\n")
            if all_ok:
                self.terminal.append_output("All required modules are installed! ✅\n")
            else:
                self.terminal.append_output(
                    "Some modules are missing. Install them via PowerShell:\n"
                    "  Install-Module <ModuleName> -Scope CurrentUser -Force\n"
                )

        self.after(0, _update)

    # ──────────────────────────── GITHUB SYNC ────────────────────────────

    def _sync_scripts(self):
        """Sync scripts from GitHub."""
        self.terminal.clear_output()
        self.sync_btn.configure(state="disabled", text="🔄  Syncing...")

        def _on_progress(msg):
            self.terminal.append_output(msg)

        def _on_done(success, msg):
            def _update():
                self.terminal.append_output(f"\n{msg}\n")
                self.sync_btn.configure(state="normal", text="🔄  Sync Scripts")
                # Refresh the current category view if one is selected
                if self._selected_category:
                    self._select_category(self._selected_category)
            self.after(0, _update)

        self.syncer.sync(progress_callback=_on_progress, done_callback=_on_done)

    # ──────────────────────────── UPDATE CHECK ────────────────────────────

    def _on_update_check(self, has_update, latest_tag, download_url, message):
        """Called when the update check completes."""
        if not has_update:
            return

        def _show():
            self.update_banner.configure(text=f"⬆ Update available: {latest_tag}")
            self.update_banner.pack(fill="x", padx=12, pady=(6, 0), before=self.sidebar.winfo_children()[2])
            self.terminal.append_output(f"\n{message}\nClick the banner in the sidebar to download.\n")
        self.after(0, _show)
