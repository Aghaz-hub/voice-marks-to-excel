# Voice Marks → Excel

Dictate marks while checking copies → Excel with auto-sum.

## Windows 10 / 11

1. Open the app in **Chrome** or **Edge**
2. **Click inside the white text box**
3. Press **Win + H**
4. Allow microphone if asked
5. Speak, e.g. `Name Rahul roll number 2K23CSUN001 question 1 A 5 marks`
6. Wait until you **see the words in the box**
7. Click **Parse & Add Student**

If Win+H does nothing:
- Settings → Privacy → Speech → On
- Settings → Time & Language → Speech → English speech pack installed
- Use a working microphone (built-in or headset)

## Android

1. Tap **inside the white text box**
2. Tap **microphone on Gboard** (not only “Browser mic”)
3. Speak until text appears in the box
4. **Parse & Add Student**

## iPhone / Mac

- iPhone: box → iOS keyboard mic → speak → Parse  
- Mac: box → Fn Fn → speak → Parse  

## Important

Grey example text is only a placeholder.  
**Parse works only when real text is visible in the white box.**

## Deploy

```bash
pip install -r requirements.txt
streamlit run app.py
```
