# GeneVennDiagram

A standalone desktop app for visualizing gene list overlaps with Venn diagrams.  
Designed for ease of use in research workflows — no coding required.  

Developed by **Wayne A. Ayers-Creech**.

---

## 🚀 Features
- Upload Excel files with gene lists (multiple sheets supported).  
- Automatically generate **symmetrical Venn diagrams**.  
- Customize colors, labels, and outputs.  
- Export results (images + Excel with overlap/unique lists).  
- Cross-platform: works on **Windows (.exe)** 

---

## 📥 Installation & Use

### Option 1: Download Executable (Recommended)
- Go to the [Releases](../../releases) page.  
- Download the latest version for your operating system:
  - `GeneVenn.exe` (Windows)  
  

No installation required — just double-click the app.  

### Option 2: Run from Source
If you prefer to run directly from Python:
```bash
git clone https://github.com/Wayne-Ayers-Creech/GeneVenn.git
cd GeneVenn
pip install -r requirements.txt
python Venn_app.py
