# CLAUDE.md

## Repo Outcome
Asset-command module for Google Sheets and Apps Script that tracks assets, driver compliance, maintenance, and fuel events with a dashboard and guided data entry.

## Operating Rules
- Preserve synthetic sample data only; never add real client names, emails, phone numbers, or deal data.
- Keep contact references aligned to `jacob@truenorthstrategyops.com` and `719-204-6365`.
- Do not commit secrets. `.env`, credential files, keys, and local machine paths are never allowed.
- Keep this repo deployable through Apps Script file import and the local `setup.js` wizard.

## Definition of Done
- README header matches TNDS standard block.
- LICENSE is MIT with TNDS copyright line.
- `.gitignore` excludes secrets, env files, virtual environments, and dependency folders.
- No hardcoded local paths or non-approved contact data in tracked source/docs.
