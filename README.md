
# 🕒 Floor Attendance Sync (2nd & 3rd Floor)

A Flask-based web app to upload, process, merge, and analyze punch-in/out Excel data for 2nd and 3rd floor attendance logs. It calculates total hours spent, applies formatting, remarks.

---

## 🚀 Features

- Upload punch-in/out logs from 2nd and 3rd floor in Excel format
- Auto-calculate total time spent by employees
- Merge both floor logs by employee and date
- Add **Remark** column based on total time compared to 8 hours
- Apply Excel-style **conditional formatting** (Green, Red, Gray)
- Include raw data logs in separate sheets
  

---

## 🛠️ How to Run Locally

### 1. **Clone the Repository**
```bash
git clone https://github.com/your-username/floor-attendance-sync.git
cd floor-attendance-sync
```

### 2. **Create a Virtual Environment (Optional but Recommended)**
```bash
python -m venv venv
source venv/bin/activate     # On Windows: venv\Scripts\activate
```

### 3. **Install Dependencies**
```bash
pip install -r requirements.txt
```

If you don’t have a `requirements.txt` yet, create one with:
```bash
pip freeze > requirements.txt
```

Or manually install main packages:
```bash
pip install flask pandas openpyxl werkzeug
```

### 4. **Run the Flask App**
```bash
python app.py
```

The app will be available at:
```
http://127.0.0.1:5000/
```

---

## 📂 Folder Structure

```
.
├── app.py                # Main Flask application
├── templates/            # HTML templates (hello.html, merged.html, upload.html, etc.)
├── uploads/              # Temporarily uploaded input files
├── outputs/              # Generated output Excel files
├── static/               # Static files (optional)
├── README.md
└── requirements.txt
```

---


---

## 📄 Pages Summary

| URL                      | Description                                 |
|--------------------------|---------------------------------------------|
| `/`                      | Home page                                   |
| `/2-3`                   | Upload and merge 2nd & 3rd floor Excel files|
| `/merged`                | Final processing with formatting + remarks  |
| `/3rdcal`                | Calculate total time from punch string      |

---

## ✅ Input File Format

### 2nd Floor Excel (Required Columns)
```
- Employee Code
- Employee Name
- Date
- Door Name.
- Time
```

### 3rd Floor Excel (Required Columns)
```
- Employee Code
- Employee Name
- Date
- Punch Records
```

---

## 📤 Output Excel

Final output includes:

- Total Time Spent (2nd + 3rd Floor)
- **Remark** column (time difference from 8 hours)
- Conditional Formatting (Green for >= 8 hrs, Red for < 8 hrs, Gray for 0)
- Additional sheets:
  - 2nd floor raw punch records (`Door Name Inputs`)
  - 3rd floor raw punch logs (`Raw Punch Records`)

---
