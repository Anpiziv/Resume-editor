# Resume Tailoring Agent

This tool edits you resume, optimizing your most relevant skills for a particular job description and requirements by using Gemini. It sells you better without making any false information about your employment history and experience. It's using a separate verified employment-history file as the source of truth.

## Setup

1. Create and activate a virtual environment.

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```
2. Install dependencies:

```cmd
python -m pip install -r requirements.txt
```

3. Set one Gemini API key environment variable in your shell:

```cmd
set /p GEMINI_API_KEY=Gemini API key: 
```

4. Add an employment_history.json template (see "employment_history_example.json" file).

5. Add a job_input.txt file (used in case the job description scraping fails) 

6. Place your private `master_resume.docx` in the project folder.

7. Edit `gemini_system_prompt.txt` if you want to adjust the Gemini writing rules without changing Python code.

## Run

```cmd
python main.py
```

Generated resumes, private resume files, job descriptions, local keys, and virtual environments are intentionally ignored by Git.

## Safety Notes

- Do not commit `.env`, API key text files, resumes, tailored outputs, or personal employment-history data.
- Keep secrets in environment variables, not in source files.
- Review generated resume changes before sharing them.
