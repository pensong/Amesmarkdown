from __future__ import annotations

from pathlib import Path
import threading

from mdconvert_app.service import convert_path


def main() -> None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "Tkinter is not available in this Python build. Install a Python distribution with Tk support, "
            "or run the CLI via `amesmarkdown`."
        ) from exc

    root = tk.Tk()
    root.title("Amesmarkdown")
    root.tk.call("tk", "appname", "Amesmarkdown")
    root.geometry("760x520")

    source_var = tk.StringVar()
    output_var = tk.StringVar()

    def append_log(message: str) -> None:
        log.configure(state="normal")
        log.insert("end", message + "\n")
        log.see("end")
        log.configure(state="disabled")

    def choose_file() -> None:
        path = filedialog.askopenfilename(
            title="Choose an Office file",
            filetypes=[
                ("Office files", "*.docx *.pptx *.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            source_var.set(path)
            if not output_var.get():
                output_var.set(str(Path(path).with_suffix("").parent / "markdown-output"))

    def choose_folder() -> None:
        path = filedialog.askdirectory(title="Choose a folder with Office files")
        if path:
            source_var.set(path)
            if not output_var.get():
                output_var.set(str(Path(path) / "markdown-output"))

    def choose_output() -> None:
        path = filedialog.askdirectory(title="Choose output folder")
        if path:
            output_var.set(path)

    def preview_selected() -> None:
        source_text = source_var.get().strip()
        if not source_text:
            messagebox.showerror("Missing source", "Choose a file or folder first.")
            return
        source = Path(source_text)
        if source.is_dir():
            messagebox.showinfo("Folder selected", "Preview is only available for single files.")
            return
        destination = Path(output_var.get().strip() or source.parent / "markdown-output")
        try:
            result = convert_path(source, destination)[0]
        except Exception as exc:  # pragma: no cover - GUI path
            messagebox.showerror("Preview failed", str(exc))
            append_log(f"Preview failed: {exc}")
            return
        preview.configure(state="normal")
        preview.delete("1.0", "end")
        preview.insert("1.0", result.markdown)
        preview.configure(state="disabled")
        append_log(f"Preview refreshed for {source.name}")

    def do_convert() -> None:
        source_text = source_var.get().strip()
        output_text = output_var.get().strip()
        if not source_text or not output_text:
            messagebox.showerror("Missing input", "Choose both a source and an output folder.")
            return

        convert_button.configure(state="disabled")
        append_log("Starting conversion...")

        def worker() -> None:
            try:
                results = convert_path(Path(source_text), Path(output_text))
            except Exception as exc:  # pragma: no cover - GUI path
                root.after(0, lambda: append_log(f"Conversion failed: {exc}"))
                root.after(0, lambda: messagebox.showerror("Conversion failed", str(exc)))
            else:
                def complete() -> None:
                    for item in results:
                        append_log(f"Converted {item.source.name} -> {item.destination}")
                    if results:
                        preview.configure(state="normal")
                        preview.delete("1.0", "end")
                        preview.insert("1.0", results[0].markdown)
                        preview.configure(state="disabled")
                    append_log("Conversion finished.")
                root.after(0, complete)
            finally:
                root.after(0, lambda: convert_button.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    root.columnconfigure(1, weight=1)
    root.rowconfigure(4, weight=1)
    root.rowconfigure(6, weight=1)

    tk.Label(root, text="Source").grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")
    tk.Entry(root, textvariable=source_var).grid(row=0, column=1, padx=12, pady=(12, 6), sticky="ew")
    tk.Button(root, text="Choose File", command=choose_file).grid(row=0, column=2, padx=12, pady=(12, 6))
    tk.Button(root, text="Choose Folder", command=choose_folder).grid(row=0, column=3, padx=(0, 12), pady=(12, 6))

    tk.Label(root, text="Output Folder").grid(row=1, column=0, padx=12, pady=6, sticky="w")
    tk.Entry(root, textvariable=output_var).grid(row=1, column=1, padx=12, pady=6, sticky="ew")
    tk.Button(root, text="Choose Output", command=choose_output).grid(row=1, column=2, padx=12, pady=6)

    button_row = tk.Frame(root)
    button_row.grid(row=2, column=0, columnspan=4, padx=12, pady=8, sticky="w")
    tk.Button(button_row, text="Preview", command=preview_selected).pack(side="left")
    convert_button = tk.Button(button_row, text="Convert", command=do_convert)
    convert_button.pack(side="left", padx=8)

    tk.Label(
        root,
        text="Drag and drop can be added later with tkinterdnd2; this version keeps the GUI dependency-light.",
        anchor="w",
    ).grid(row=3, column=0, columnspan=4, padx=12, pady=(0, 8), sticky="ew")

    tk.Label(root, text="Preview").grid(row=4, column=0, padx=12, sticky="nw")
    preview = scrolledtext.ScrolledText(root, wrap="word", state="disabled", height=12)
    preview.grid(row=4, column=1, columnspan=3, padx=12, pady=(0, 12), sticky="nsew")

    tk.Label(root, text="Status").grid(row=6, column=0, padx=12, sticky="nw")
    log = scrolledtext.ScrolledText(root, wrap="word", state="disabled", height=8)
    log.grid(row=6, column=1, columnspan=3, padx=12, pady=(0, 12), sticky="nsew")

    root.mainloop()


if __name__ == "__main__":
    main()
