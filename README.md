# My Free Speech Reader

A high-productivity text-to-speech (TTS) application built in PowerShell. Designed for professionals who need to consume long-form text or screen content without the bulk of a traditional GUI.

## 🌟 Key Features
- **Draggable Mini-Widget:** Toggle between a full settings dashboard and a compact, borderless widget that stays on top of other windows.
- **Global Hotkeys:** Control playback from any application using `Ctrl + Alt + Shift` shortcuts.
- **OCR Integration:** Leverages the Tesseract engine to extract and read text from images or non-selectable screen areas.
- **Dual Voice Support:** Switch between local SAPI5 voices and online natural voices.
- **System Tray Integration:** Runs quietly in the background to save taskbar space.

## ⌨️ Global Shortcuts
- **Read from Screen:** `Ctrl + Alt + Shift + R`
- **Pause/Resume:** `Ctrl + Alt + Shift + P` / `U`
- **Next/Previous Paragraph:** `Ctrl + Alt + Shift + N` / `B`
- **Stop:** `Ctrl + Alt + Shift + S`

## 🛠️ Installation
1. Navigate to the **Releases** section on the right.
2. Download the latest `setup.exe`.
3. Run the installer (this will automatically bundle the necessary `Tesseract.dll` and language data).
4. Launch the app from your Desktop or Start Menu.

## 📂 Project Structure
- `My Free Speech.ps1`: The core logic and WinForms GUI.
- `Tesseract.dll`: OCR library for image-to-text processing.
- `tessdata/`: Language training data for the OCR engine.

---
*Developed as a utility for accessibility and productivity.*
