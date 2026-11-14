# OneDrive Text

A small web app for opening and editing plain-text files in OneDrive from an iPhone or iPad.

Live: https://mark-onedrive-text.netlify.app

## Why it exists

Saving through the iOS share sheet into OneDrive never overwrites a file; the OneDrive file provider adds a number to the name instead. This app talks to OneDrive directly through Microsoft Graph, so Save asks "Replace?" and then overwrites the original file in place.

## Features

- Sign in with a Microsoft account and browse or search OneDrive folders
- Save asks before replacing, and warns if the file changed elsewhere since it was opened
- Renaming in the title box saves a new file in the same folder (asks if that name is taken)
- Keeps the file's line endings and byte-order mark as they were
- Unsaved edits are kept on the device as a draft
- Find and replace, mono or serif type, text size, line wrap
- iPad keyboard shortcuts: Cmd-S save, Cmd-F find
- Add to Home Screen to use it like an app

## Setup

The app uses a Microsoft Entra app registration (a public single-page app, no secret):

- Account types: **Accounts in any organizational directory and personal Microsoft accounts** (`signInAudience: AzureADandPersonalMicrosoftAccount`, `requestedAccessTokenVersion: 2`)
- Redirect URI: **Single-page application**, `https://mark-onedrive-text.netlify.app/`
- Delegated permissions: `User.Read`, `Files.ReadWrite`

The client ID is set in `BUILT_IN_CLIENT_ID` in `index.html`. If it is left empty, the app shows a setup screen where a client ID can be pasted in instead.

## Files

- `index.html`: the whole app
- `msal-browser-5.25.1.min.js`: Microsoft's sign-in library (@azure/msal-browser), served locally
- `manifest.webmanifest`, icons: Home Screen app metadata
- `_headers`, `netlify.toml`: Netlify hosting config (static site, no build step)
