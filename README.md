# RSS to OneNote Sync

This project automates syncing RSS feeds to OneNote using Python and GitHub Actions.

## Features
- Fetches content from 24 RSS feeds.
- Cleans and processes HTML content (removes scripts, ads, and links).
- Creates formatted OneNote pages with date, source, and original link.
- Runs daily at 2:00, 10:00, 14:00, 18:00 (Beijing time).

## Setup
- Configure `AZURE_CLIENT_ID` in GitHub Secrets.
- Run manually to complete device code authentication.
- Key files: `rss_to_onenote.py`, `requirements.txt`, `.github/workflows/sync.yml`.
