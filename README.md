# Resume Tailoring Agent

Python tool that tailors a private master resume to a target job description with Gemini, using a separate verified employment-history file as the source of truth.

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

3. Set one Gemini API key environment variable in your shell:

```powershell
$env:GEMINI_API_KEY = Read-Host "Gemini API key"
```

4. Copy the private input templates:

```powershell
Copy-Item employment_history.example.json employment_history.json
Copy-Item job_input.example.txt job_input.txt
```

5. Place your private `master_resume.docx` in the project folder.

6. Edit `gemini_system_prompt.txt` if you want to adjust the Gemini writing rules without changing Python code.

## Run

```powershell
python main.py
```

Generated resumes, private resume files, job descriptions, local keys, and virtual environments are intentionally ignored by Git.

## Safety Notes

- Do not commit `.env`, API key text files, resumes, tailored outputs, or personal employment-history data.
- Keep secrets in environment variables, not in source files.
- Review generated resume changes before sharing them.
