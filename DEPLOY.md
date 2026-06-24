# Deploy — Truck Plan Builder (single Flask app)

This is the simplest layout Vercel's Python runtime accepts. Just THREE files:

```
app.py            everything: the engine + the web page + /api/process + /welcome.png
requirements.txt  flask, openpyxl, xlrd
welcome.png       (optional) your header image — add it next to app.py
```

There is NO index.html, NO api/ folder, and NO vercel.json. Vercel sees `app.py`
(a Flask `app`) at the root and runs the whole site from it.

## Steps (GitHub + Vercel)
1. In your GitHub repo, make sure the ONLY files are `app.py`, `requirements.txt`,
   `.gitignore` (and later `welcome.png`). **Delete** any old `index.html`,
   `api/` folder, and `vercel.json` from the repo.
2. Commit. Vercel redeploys automatically (or Add New -> Project -> Import).
3. Open the URL: the page loads at `/`, uploads post to `/api/process`.

## Add the welcome image
Upload your picture to the repo root named exactly `welcome.png`, commit. It shows
in the "Welcome DRR!" header (served by app.py at /welcome.png). If absent, the
header just shows the text.

## Why this layout
Vercel's current Python builder wants a single app entrypoint at a standard path
(app.py / index.py / etc.) exposing a Flask `app`. Per-file `api/*.py` functions and
the `handler` class are no longer the happy path. One `app.py` that serves both the
page and the API sidesteps every error we hit (functions pattern, handler-vs-app,
invalid vercel.json, missing entrypoint).

## Local test
    pip install flask openpyxl xlrd
    python app.py        # then open http://localhost:8000  (set app.run port if needed)
(For local runs you can add at the bottom of app.py:
    if __name__ == "__main__": app.run(port=8000, debug=True)
This block is ignored by Vercel.)
