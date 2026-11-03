# Security Policy

## Supported versions

Only the latest commit on the `main` branch is supported.

## Reporting a vulnerability

Please do not open a public issue for security problems.

Use GitHub's private vulnerability reporting: open the **Security** tab of this repository and choose **Report a vulnerability**. Include the affected component, reproduction steps and the impact you observed.

You can expect an acknowledgement within a few days. Confirmed issues are fixed on `main` and disclosed after a fix is available.

## Scope notes

- The app reads Word documents and writes to the current user's Outlook signatures folder; conversion guards against path traversal in asset paths.
- Dependencies are monitored with Dependabot and secret scanning with push protection.
