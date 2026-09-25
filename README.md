This is an app to integrate multiple .xlsx files into one file, and organize the information, and make summary. It uses an AI model, Jev, which was made by TypeSafe AI, when the problem cannot be solved by rules. Now, the app is compatible only for .xlsx files written in Japanese, French, and English.

## Result of the final run

```
36 files, 5 offices, 3 languages  ->  977 rows

861 clean            (88%)
 18 decided by Jev    (2%)
 98 left for a person (10%)
```

## HOW TO USE

```powershell

.venv\Scripts\python.exe -m pip install -r requirements.txt

copy .env.example .env ##please get your own!

.venv\Scripts\python.exe src\app.py
```

The result is written to `output/master.xlsx`.

## Folder structure

```
.
├── src/
│   ├── app.py
│   └── History/
├── settings.xlsx         The offices, and the conventions of their countries
├── input/                36 inquiry files: 5 offices, 3 languages
├── output/
│   ├── master.xlsx       The result: Summary, Data, Issues, Issues detail
│   └── 5.x_*.xlsx
├── tests/
├── answer_key/
├── tools/
├── requirements.txt      Pinned versions. Do not upgrade numpy!!
└── .env.example      	Rename it to .env and get paste your own JEV API key!
```
## Notes

- **The file name decides the office.** `inquiries_paris_2026-03.xlsx` is read as
  the Paris office. A file whose name matches no office in `settings.xlsx` is not
  read; the app records the reason and continues.
- **Do not upgrade numpy.** On my machine the newest version was blocked by
  Windows application control, so `requirements.txt` pins 2.3.5.
- **The app runs without an API key.** It does not stop. Every question that
  would have gone to the model is recorded and sent to a person instead.
