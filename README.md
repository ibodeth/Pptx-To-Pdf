# 📄 PPTX to PDF Converter (Offline, Clean & Media-Aware)

**Convert pptx to pdf without internet and junk apps.**

This project allows you to **batch convert PowerPoint files (`.pptx`) to PDF** using **LibreOffice** in headless mode — completely **offline**, without ads, trackers, online services, or unnecessary third‑party applications.

In addition, it can **optionally remove embedded media (videos / audio)** from presentations **before conversion**, reducing file size and avoiding heavy PDFs.


## 🚀 Features

* ✅ 100% **Offline** (no internet required)
* 🧹 No ads, no trackers, no online converters
* 📂 Batch conversion of all `.pptx` files in a folder
* 🎞️ **Optional media (video/audio) removal** before PDF export
* 📉 Smaller PDF sizes when media removal is enabled
* 🖥️ Uses **LibreOffice headless CLI**
* 🎛️ Simple GUI (Tkinter dialogs)
* 📄 PDFs are saved in the **same folder** as originals
* 📊 Clear success / error reporting

> ⚠️ `.ppt` files are supported for conversion, but **media removal works only with `.pptx`**.

---

## 🛠 Technologies Used

* Python
* Tkinter (built‑in GUI)
* LibreOffice (`soffice` CLI)
* `python-pptx`
* `subprocess`, `os`

---

## 📋 Requirements

### ✅ Required Software

* **LibreOffice** (installed locally)
* Windows 10 / 11

**Default LibreOffice path used in the script:**

```
C:\Program Files\LibreOffice\program\soffice.exe
```

If LibreOffice is installed elsewhere, update the `LIBREOFFICE_PATH` variable in the script.

---

## 📥 Installation

### 1️⃣ Clone the Repository

```bash
git clone https://github.com/ibodeth/Pptx-To-Pdf.git
cd Pptx-To-Pdf
```

### 2️⃣ (Optional) Create a Virtual Environment

```bash
python -m venv venv
venv\Scripts\activate
```

---

## 📦 Dependencies

The only external dependency is **python-pptx** (used for media removal).

Install it with:

```bash
pip install python-pptx
```

If the package is missing, the script will notify you at startup.

---

## ▶️ Usage

Run the script:

```bash
python pptx_to_pdf.py
```

### Workflow

1. Select the folder containing `.pptx` files
2. Choose whether to **remove embedded videos/audio**
3. Files are processed one by one
4. PDFs are created in the same folder
5. Temporary cleaned files are automatically deleted

---

## 🎞️ Media Removal Mode (Optional)

When enabled:

* All **video and audio shapes** are removed from slides
* Presentations are saved temporarily without media
* PDFs are generated from the cleaned files
* Temporary files are deleted automatically

✅ Result: **smaller, lighter PDFs**

---

## 📂 Example Directory Structure

```
Presentations/
├── demo.pptx
├── lecture_with_video.pptx
├── demo.pdf
├── lecture_with_video.pdf
```

---

## ⚠️ Common Issues & Solutions

### LibreOffice Not Found

```
LibreOffice was not found!
```

✅ **Fix:** Update the `LIBREOFFICE_PATH` variable to match your LibreOffice installation path.

---

### python-pptx Not Installed

```
Please run: pip install python-pptx
```

✅ **Fix:**

```bash
pip install python-pptx
```

---

## 🧠 How It Works (Internals)

* Tkinter is used for folder selection and option dialogs
* `.pptx` files are scanned
* Optional media removal via `python-pptx`
* LibreOffice runs in `--headless` mode to export PDFs
* Temporary files are cleaned automatically

---

## 👨‍💻 Author

**İbrahim Nuryağınlı**

* GitHub: [https://github.com/ibodeth](https://github.com/ibodeth)
* YouTube: [https://www.youtube.com/@ibrahim.python](https://www.youtube.com/@ibrahim.python)
* LinkedIn: [https://www.linkedin.com/in/ibrahimnuryaginli/](https://www.linkedin.com/in/ibrahimnuryaginli/)

---

## 📄 License

This project is licensed under the **MIT License**.

⭐ If you find this project useful, please consider giving it a star on GitHub!
