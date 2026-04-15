# Amesmarkdown macOS Launcher

This folder contains a minimal macOS `.app` launcher bundle for Amesmarkdown.

## What it does

- Shows up as `Amesmarkdown.app`
- Launches the existing Python GUI
- Avoids needing to type a Terminal command every time

## Expected project location

The launcher expects the project folder to live at:

```text
/Users/pensong/Documents/Amesmarkdown
```

It also expects the Python virtual environment at:

```text
/Users/pensong/Documents/Amesmarkdown/.venv
```

## Install steps

1. Rename the project folder to `Amesmarkdown`
2. Reinstall the project into its venv
3. Copy `macos/Amesmarkdown.app` into `/Applications`

## Reinstall commands after rename

```bash
cd /Users/pensong/Documents/Amesmarkdown
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -e .
```
