# D365 Customer Service Copilot Test

Playwright-based test that sends each prompt from `Prompts and Responses.xlsx` to the D365 Customer Service Copilot, captures the response, compares it against the expected answer, and writes results to a new Excel file.

## Prerequisites

- **Node.js** 18+ installed

## Setup

1. **Install dependencies:**
   ```
   npm install
   ```

2. **Install the Playwright Chromium browser:**
   ```
   npx playwright install chromium
   ```

That's it — authentication and the D365 URL are handled interactively on first run.

## Running the Test

```
npx playwright test
```

On first run you'll see an interactive setup:

```
╔══════════════════════════════════════════╗
║         D365 Copilot Test Setup          ║
╚══════════════════════════════════════════╝

  Current D365 URL: https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com
  Enter new URL or press Enter to keep:
```

- **D365 URL** — paste your org's URL (e.g. `https://yourorg.crm.dynamics.com`), or press Enter to keep the saved value.
- **Login** — if no saved session exists, a Chromium browser opens for you to sign in to D365. Once logged in, press Enter in the terminal. Your session is saved to `auth-state.json` and reused on future runs.

To force a fresh login, delete `auth-state.json` and run the test again.

### If Chromium is not installed

The script will exit with a clear message:
```
  ┌──────────────────────────────────────────────────────┐
  │  Chromium is not installed!                          │
  │                                                      │
  │  Run this command first:                             │
  │    npx playwright install chromium                   │
  └──────────────────────────────────────────────────────┘
```

## Customising selectors

The script includes default selectors for the D365 Copilot side-panel chat. If your environment differs, update the selectors in `tests/d365-copilot-test.spec.ts` — look for comments marked `// Update these selectors`.

## Configuration (.env overrides)

The interactive prompts handle most settings, but you can also set these in `.env` to skip prompts or tune behaviour:

| Variable | Default | Description |
|---|---|---|
| `D365_URL` | *(prompted)* | Pre-fills the URL prompt on first run |
| `COPILOT_RESPONSE_TIMEOUT` | `60` | Seconds to wait for each Copilot response |
| `SIMILARITY_THRESHOLD` | `0.6` | 0–1 word-overlap threshold for pass/fail |

## Output

Results are saved to `Test Results YYYY-MM-DD.xlsx` with these columns:

| Column | Description |
|---|---|
| **#** | Prompt number |
| **Prompt** | The question sent to Copilot |
| **Expected Response** | From the source spreadsheet |
| **Actual Response** | What Copilot returned |
| **Similarity** | Word-overlap percentage |
| **Result** | PASS (green) or FAIL (red) |
| **Referenced Docs** | From the source spreadsheet |

A summary row at the bottom shows total pass/fail counts.
