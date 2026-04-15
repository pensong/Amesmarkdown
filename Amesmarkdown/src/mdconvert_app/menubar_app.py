from __future__ import annotations

from pathlib import Path
import subprocess
import threading
from typing import Optional

from mdconvert_app.service import convert_path


def main() -> None:
    try:
        import rumps
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "rumps is not available in this Python build. Install project dependencies and retry `amesmarkdown-menubar`."
        ) from exc

    class MenuBarConverter(rumps.App):
        def __init__(self) -> None:
            super().__init__("AM", quit_button=None)
            self.menu = [
                rumps.MenuItem("Convert File...", callback=self.convert_file),
                rumps.MenuItem("Convert Folder...", callback=self.convert_folder),
                rumps.MenuItem("Open Output Folder", callback=self.open_output_folder),
                None,
                rumps.MenuItem("Quit", callback=self.quit_app),
            ]
            self.last_output_folder: Optional[Path] = None

        def convert_file(self, _: object) -> None:
            source_path = self._choose_file()
            if not source_path:
                return
            output_dir = self._choose_output_dir(default=source_path.with_suffix("").parent / "markdown-output")
            if not output_dir:
                return
            self._start_conversion(source_path, output_dir)

        def convert_folder(self, _: object) -> None:
            source_dir = self._choose_directory(title="Choose a folder with Office files")
            if not source_dir:
                return
            output_dir = self._choose_output_dir(default=source_dir / "markdown-output")
            if not output_dir:
                return
            self._start_conversion(source_dir, output_dir)

        def open_output_folder(self, _: object) -> None:
            if not self.last_output_folder:
                rumps.alert("No output folder yet", "Run a conversion first so there is a folder to open.")
                return
            subprocess.run(["open", str(self.last_output_folder)], check=False)

        def quit_app(self, _: object) -> None:
            rumps.quit_application()

        def _start_conversion(self, source: Path, output_dir: Path) -> None:
            self.title = "AM*"
            self.last_output_folder = output_dir
            rumps.notification("Amesmarkdown", "Conversion started", f"{source.name} -> {output_dir}")

            def worker() -> None:
                try:
                    results = convert_path(source, output_dir)
                except Exception as exc:
                    self.title = "AM"
                    rumps.notification("Amesmarkdown", "Conversion failed", str(exc))
                else:
                    self.title = "AM"
                    if results:
                        rumps.notification(
                            "Amesmarkdown",
                            "Conversion finished",
                            f"Created {len(results)} Markdown file(s) in {output_dir}",
                        )
                    else:
                        rumps.notification(
                            "Amesmarkdown",
                            "No supported files found",
                            f"No .docx, .pptx, or .xlsx files were found in {source}",
                        )

            threading.Thread(target=worker, daemon=True).start()

        def _choose_file(self) -> Optional[Path]:
            return self._run_file_dialog(
                chooser=self._ask_open_file,
                title="Choose an Office file",
                initial_dir=None,
            )

        def _choose_output_dir(self, default: Path) -> Optional[Path]:
            selected = self._run_file_dialog(
                chooser=self._ask_directory,
                title="Choose output folder",
                initial_dir=default.parent,
            )
            return selected or default

        def _choose_directory(self, title: str) -> Optional[Path]:
            return self._run_file_dialog(
                chooser=self._ask_directory,
                title=title,
                initial_dir=None,
            )

        def _run_file_dialog(self, chooser, title: str, initial_dir: Optional[Path]) -> Optional[Path]:
            try:
                import tkinter as tk
                from tkinter import filedialog
            except ModuleNotFoundError as exc:
                raise SystemExit(
                    "Tkinter is required for file/folder pickers in menu bar mode. Use a Python build with Tk support."
                ) from exc

            root = tk.Tk()
            root.withdraw()
            root.update()
            try:
                selected = chooser(filedialog, title, initial_dir)
            finally:
                root.destroy()
            if not selected:
                return None
            return Path(selected).expanduser().resolve()

        @staticmethod
        def _ask_open_file(filedialog, title: str, initial_dir: Optional[Path]) -> str:
            options = {
                "title": title,
                "filetypes": [
                    ("Office files", "*.docx *.pptx *.xlsx"),
                    ("All files", "*.*"),
                ],
            }
            if initial_dir:
                options["initialdir"] = str(initial_dir)
            return filedialog.askopenfilename(**options)

        @staticmethod
        def _ask_directory(filedialog, title: str, initial_dir: Optional[Path]) -> str:
            options = {"title": title}
            if initial_dir:
                options["initialdir"] = str(initial_dir)
            return filedialog.askdirectory(**options)

    MenuBarConverter().run()


if __name__ == "__main__":
    main()
