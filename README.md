# Voice Marks → Excel (Faculty Dictation App)

Faculty checks answer copies while speaking marks into a **wired earphone mic** connected to an **Android phone**.  
The app converts speech → structured marks → **Excel with auto-sum formulas**.

---

## How it works (Android)

1. Open the Streamlit app on the phone browser  
2. Tap the text box  
3. Tap the **microphone on Gboard** (Google voice typing)  
4. Speak while checking, for example:

```
Name Rahul Sharma
Roll number 2K23CSUN001
Question 1 A 5 marks
Question 1 B 3 marks
Question 2 zero
Question 5 4 marks
Question 7 2 marks
```

5. Tap **Parse & Add Student**  
6. Repeat for the next student  
7. Download Excel when finished

No extra app or special hardware beyond a normal wired earphone with mic.

---

## What the Excel contains

| Sheet | Content |
|-------|---------|
| **Summary** | Name, Roll Number, Grand Total (one row per student) |
| **Student Marks (Flat)** | Name, Roll, 1(a), 1(b)… Q totals, Grand Total — all with Excel formulas |
| **Info** | Generation details |

---

## Deploy on Streamlit Cloud

1. Create a new GitHub repository  
2. Upload all files from this folder  
3. Go to [https://share.streamlit.io](https://share.streamlit.io) → New app → `app.py` → Deploy  

No system packages required (no Tesseract / Poppler).

---

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Speaking tips

- Say **“question 1 A 5 marks”** clearly  
- Use **“zero”** for 0  
- One student per dictation is most reliable  
- You can edit the text in the box before tapping Parse  

---

Made for faculty who want hands-free marks entry while checking copies.
