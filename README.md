# PDF Tools Pro

**Free • Private • 100% Client-Side PDF Utility App**  
Built for **M/s Sushila Engineering Works**

---

## Features

| Tool              | Description                                      |
|-------------------|--------------------------------------------------|
| PDF Viewer        | High-quality multi-page PDF viewer               |
| Photo to PDF      | Convert JPG / PNG images into a PDF              |
| Merge PDF         | Combine multiple PDF files into one              |
| Split PDF         | Extract each page as a separate PDF              |
| Compress PDF      | Reduce file size (object streams)                |
| PDF to Word       | Extract text → editable `.docx`                  |
| Word / Text to PDF| Convert text into a clean PDF                    |
| Scan & Sign       | Camera scan + digital signature pad              |

All processing happens **inside your browser / phone**. No files are uploaded to any server.

---

## How to get the APK (GitHub Actions)

1. Create a new **public** GitHub repository.
2. Upload **all files** from this ZIP (keep the folder structure).
3. Go to the **Actions** tab.
4. Click **Build Android APK** on the left.
5. Click the **Run workflow** button → **Run workflow**.
6. Wait 2–4 minutes until the green checkmark appears.
7. Click the completed run → scroll down to **Artifacts**.
8. Download **PDF-Tools-Pro-APK**.
9. Transfer `app-debug.apk` to your Motorola phone and install it.

---

## Local testing (optional)

```bash
npm install
npx serve www
```

Open the URL shown in the terminal.

---

## Build APK locally (optional)

```bash
npm install
npx cap add android
npx cap sync
npx cap open android
```

Then build from Android Studio.

---

© 2026 M/s Sushila Engineering Works
