# Permanent Online Deployment

Recommended host: Streamlit Community Cloud.

1. Upload this project to GitHub.
2. Open https://share.streamlit.io/ and sign in with GitHub.
3. Create a new app from your GitHub repository.
4. Set the main file path to:

```text
app_files/app.py
```

5. Keep the dependency file as:

```text
app_files/requirements.txt
```

6. Choose the app URL/subdomain you want, for example:

```text
roger-campbell-pm
```

7. Deploy.

Notes:
- The app reads its data, assets, sources, and configuration relative to `app_files/app.py`.
- Do not upload `.venv`, `__pycache__`, log files, or tunnel files.
- The Windows launcher is only for local running. The online host should run `app_files/app.py`.
