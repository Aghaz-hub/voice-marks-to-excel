# Voice Marks → Excel (Faculty Dictation App)

Speak marks while checking answer copies → **Excel** with auto-sum formulas.

Works on **iPhone, iPad, Mac, Windows, and Android**.

---

## iPhone / iPad

1. Open the app in **Safari** or **Chrome**
2. Tap the text box
3. Tap the **microphone icon** on the iOS keyboard (dictation)
4. Speak, for example:

```
Name Rahul Sharma
Roll number 2K23CSUN001
Question 1 A 5 marks
Question 1 B 3 marks
Question 2 zero
Question 5 4 marks
```

5. Tap **Parse & Add Student**
6. Repeat for the next student → **Download Excel**

Wired **EarPods** / USB-C earphones with mic work automatically.

If the keyboard mic is hidden: tap the text box and swipe the keyboard options, or enable Dictation under  
**Settings → General → Keyboard → Enable Dictation**.

---

## Mac (Apple laptop)

1. **System Settings → Keyboard → Dictation → On**
2. Open the app in **Safari** or **Chrome**
3. Click inside the text box
4. Press **Fn twice** (or the **Globe** key) to start dictation
5. Speak → click **Parse & Add Student**

**Optional:** click **Start browser mic**, allow the microphone, speak, then copy the heard text into the box and Parse.

---

## Windows laptop

1. Click inside the text box  
2. Press **Win + H**  
3. Speak → **Parse & Add Student**

---

## Android phone

1. Tap the text box → **Gboard mic**  
2. Speak → **Parse & Add Student**

---

## Excel output

| Sheet | Content |
|-------|---------|
| Summary | Name, Roll, Grand Total |
| Student Marks (Flat) | All parts + formula totals |
| Info | Generation details |

---

## Deploy

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or upload to GitHub → [share.streamlit.io](https://share.streamlit.io) → deploy `app.py`.

---

## Speaking tips

- “Question 1 A 5 marks”  
- Say **zero** for 0  
- One student per dictation  
- Edit the text before parsing if needed  
