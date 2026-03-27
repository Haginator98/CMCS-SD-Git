import subprocess
import threading
import os
import signal
from pathlib import Path

import customtkinter as ctk


class TerminalWidget(ctk.CTkFrame):
    """Embedded terminal widget for running PowerShell scripts and displaying output."""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self._process: subprocess.Popen | None = None
        self._read_thread: threading.Thread | None = None
        self._running = False

        # Output display
        self.output_text = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(family="Menlo", size=13),
            wrap="word",
            state="disabled",
            fg_color="#1e1e2e",
            text_color="#cdd6f4",
            corner_radius=8,
        )
        self.output_text.pack(fill="both", expand=True, padx=4, pady=(4, 2))

        # Input area
        self.input_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.input_frame.pack(fill="x", padx=4, pady=(2, 4))

        self.input_label = ctk.CTkLabel(
            self.input_frame, text="›", font=ctk.CTkFont(family="Menlo", size=14, weight="bold"),
            text_color="#a6e3a1",
        )
        self.input_label.pack(side="left", padx=(4, 2))

        self.input_entry = ctk.CTkEntry(
            self.input_frame,
            font=ctk.CTkFont(family="Menlo", size=13),
            placeholder_text="Type input here and press Enter...",
            fg_color="#1e1e2e",
            text_color="#cdd6f4",
            border_color="#45475a",
            corner_radius=6,
        )
        self.input_entry.pack(side="left", fill="x", expand=True, padx=(2, 4))
        self.input_entry.bind("<Return>", self._on_enter)

        # Stop button
        self.stop_btn = ctk.CTkButton(
            self.input_frame,
            text="⏹ Stop",
            width=70,
            fg_color="#f38ba8",
            hover_color="#d35f8d",
            text_color="#1e1e2e",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self.stop_script,
            corner_radius=6,
        )
        self.stop_btn.pack(side="right", padx=(4, 0))

        self._update_input_state()

    def append_output(self, text: str):
        """Append text to the output display (thread-safe via after)."""
        def _do():
            self.output_text.configure(state="normal")
            self.output_text.insert("end", text)
            self.output_text.see("end")
            self.output_text.configure(state="disabled")
        self.after(0, _do)

    def clear_output(self):
        self.output_text.configure(state="normal")
        self.output_text.delete("1.0", "end")
        self.output_text.configure(state="disabled")

    def run_script(self, script_path: Path, pwsh_path: str = "pwsh"):
        """Run a PowerShell script in a subprocess."""
        if self._running:
            self.append_output("\n⚠ A script is already running. Stop it first.\n")
            return

        self.clear_output()
        self.append_output(f"▶ Running: {script_path.name}\n{'─' * 50}\n")

        self._running = True
        self._update_input_state()

        def _run():
            try:
                self._process = subprocess.Popen(
                    [pwsh_path, "-NoProfile", "-NoLogo", "-File", str(script_path)],
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    cwd=str(script_path.parent),
                    env={**os.environ, "NO_COLOR": "1"},
                    preexec_fn=os.setsid,
                )

                for line in iter(self._process.stdout.readline, ""):
                    if not self._running:
                        break
                    self.append_output(line)

                self._process.wait()
                exit_code = self._process.returncode
                self.append_output(f"\n{'─' * 50}\n")
                if exit_code == 0:
                    self.append_output("✅ Script finished successfully.\n")
                else:
                    self.append_output(f"⚠ Script exited with code {exit_code}.\n")

            except FileNotFoundError:
                self.append_output(
                    "\n❌ PowerShell (pwsh) not found.\n"
                    "Install it with: brew install powershell/tap/powershell\n"
                )
            except Exception as e:
                self.append_output(f"\n❌ Error: {e}\n")
            finally:
                self._process = None
                self._running = False
                self.after(0, self._update_input_state)

        self._read_thread = threading.Thread(target=_run, daemon=True)
        self._read_thread.start()

    def run_command(self, command: str, pwsh_path: str = "pwsh"):
        """Run an arbitrary PowerShell command."""
        if self._running:
            self.append_output("\n⚠ A process is already running.\n")
            return

        self.clear_output()
        self.append_output(f"▶ Running command...\n{'─' * 50}\n")

        self._running = True
        self._update_input_state()

        def _run():
            try:
                self._process = subprocess.Popen(
                    [pwsh_path, "-NoProfile", "-NoLogo", "-Command", command],
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                    env={**os.environ, "NO_COLOR": "1"},
                    preexec_fn=os.setsid,
                )

                for line in iter(self._process.stdout.readline, ""):
                    if not self._running:
                        break
                    self.append_output(line)

                self._process.wait()
                self.append_output(f"\n{'─' * 50}\n✅ Done.\n")

            except Exception as e:
                self.append_output(f"\n❌ Error: {e}\n")
            finally:
                self._process = None
                self._running = False
                self.after(0, self._update_input_state)

        self._read_thread = threading.Thread(target=_run, daemon=True)
        self._read_thread.start()

    def stop_script(self):
        """Stop the running script."""
        if self._process and self._running:
            self._running = False
            try:
                os.killpg(os.getpgid(self._process.pid), signal.SIGTERM)
            except (ProcessLookupError, OSError):
                pass
            self.append_output("\n🛑 Script stopped by user.\n")
            self._update_input_state()

    def _on_enter(self, _event=None):
        """Send input to the running process."""
        if self._process and self._process.stdin and self._running:
            user_input = self.input_entry.get()
            self.input_entry.delete(0, "end")
            try:
                self._process.stdin.write(user_input + "\n")
                self._process.stdin.flush()
                self.append_output(f"» {user_input}\n")
            except (BrokenPipeError, OSError):
                self.append_output("⚠ Cannot send input — process ended.\n")

    def _update_input_state(self):
        if self._running:
            self.input_entry.configure(state="normal")
            self.stop_btn.configure(state="normal")
        else:
            self.input_entry.configure(state="disabled")
            self.stop_btn.configure(state="disabled")

    @property
    def is_running(self) -> bool:
        return self._running
