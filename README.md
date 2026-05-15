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
2. Download the latest `mysetup.exe`.
3. Run the installer (this will automatically bundle the necessary `Tesseract.dll` and language data).
4. Launch the app from your Desktop or Start Menu.

## 📂 Project Structure
- `My Free Speech.ps1`: The core logic and WinForms GUI.
- `Tesseract.dll`: OCR library for image-to-text processing.
- `tessdata/`: Language training data for the OCR engine.

## ⬇️ How to install online Natural Voices on MyFreeSpeech App
> when you Install the MyTTS app, it will use the locally available voices like microsoft David that usually do not sound natural. Inorder to get more natural sounding voices, follow the following steps
<br><br>
<img width="1366" height="768" alt="image" src="https://github.com/user-attachments/assets/c56d83aa-b7aa-47c1-a3a2-02c6d5f13a69" /><br><br>

Step 1️⃣: Navigate to the Releases page: On the GitHub repository for NaturalVoiceSAPIAdapter <href>https://github.com/gexgd0419/NaturalVoiceSAPIAdapter</href>

Step 2️⃣: look at the sidebar on the right side of the screen and click on "Releases".

Step 3️⃣: Download the latest Assets: Look for the most recent version (e.g., v1.x.x) and scroll down to the "Assets" section at the bottom of that post.

Step 4️⃣: Select the ZIP file: Download the file named something like NaturalVoiceSAPIAdapter.zip (avoid the ones that say "Source code").

Step 5️⃣: Extract the new ZIP: Once this new folder is unzipped, the Installer.exe and NaturalVoiceSAPIAdapter.exe files will be present.

Step 6️⃣: Run Installer.exe: Right-click the file and select "Run as Administrator".

Step 7️⃣: Install for Both: Click the Install button for both the 32-bit and 64-bit versions to ensure full compatibility.

Step 8️⃣: Check narrator natural voices and specify the path where you unzipped the NaturalVoiceSAPIAddapter.

Step 9️⃣: Select Edge Voices: Check the box for "Include Microsoft Edge natural voices" since the Narrator options are missing from the system settings.


> After the installer finishes, these high-quality voices are registered as standard system voices. When ScreenReader3.ps1 is launched, the Update-VoiceList function will automatically find these new "Natural" entries and add them to the dropdown menu. They will thus appear under the sapi voices on our tool.
   



---
*Developed as a utility for accessibility and productivity.*

*This project demonstrates the integration of .NET classes, Win32 API (P/Invoke) for global hooks, and OCR libraries within a PowerShell environment." This highlights your technical skills to potential employers.*
