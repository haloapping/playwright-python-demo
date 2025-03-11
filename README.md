# Description

Demo Playwright app internal Katalon team. Simple app to introduce basic features Playwright.

## How to use?

- Clone project with git.
- Install Python Intepreter [Python Website](https://www.python.org/)
- Run command `python -m venv .env` on root project.
- Activate env with command `cd .env/bin && activate` on Windows and `cd .env/bin && source activate` on Linux/MacOS
- Install all deps with command `pip install requirements.txt`.
- Run main program with command `pytest test_main.py -s --browser chromium`.
- You can run trace file after run main program with command `playwright show-trace trace.zip`.
