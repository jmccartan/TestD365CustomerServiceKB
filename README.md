# D365 Customer Service Copilot Test

Playwright-based test that sends each prompt from `Prompts and Responses.xlsx` to the D365 Customer Service Copilot, captures the response, compares it against the expected answer, and writes results to a new Excel file.

## Prerequisites

- **Node.js** 18+ installed
- **Microsoft Edge** installed (the test uses your Edge profiles for authentication)

## Setup

1. **Install dependencies:**
   ```
   npm install
   ```

2. **Install the Playwright Edge browser driver:**
   ```
   npx playwright install msedge
   ```

That's it — no manual config files to edit. Everything else is handled interactively on first run.

## Running the Test

```
npx playwright test
```

On startup you'll see an interactive setup prompt:

```
╔══════════════════════════════════════════╗
║         D365 Copilot Test Setup          ║
╚══════════════════════════════════════════╝

  Current D365 URL: https://REPLACE_WITH_YOUR_ORG.crm.dynamics.com
  Enter new URL or press Enter to keep:

  Edge profiles found:
  [1]  Personal  (Default)
  [2]  Work  (Profile 1)
  [3]  Demo - admin@contoso.onmicrosoft.com  (Profile 2) ◄ current
  [0]  Skip — use auth-state.json fallback

  Enter number or press Enter to keep default:
```

- **D365 URL** — paste your org's URL (e.g. `https://yourorg.crm.dynamics.com`), or press Enter to keep the current value.
- **Edge profile** — pick a profile that is already signed in to D365. The test copies it to a temp directory so it won't interfere with your running Edge.

Both choices are saved to `.test-settings.json` so on subsequent runs you can just press Enter twice to reuse them.

### Fallback authentication

If you don't want to use an Edge profile, select `[0]` at the profile prompt and instead run:
```
npx ts-node auth-setup.ts
```
This opens a browser for you to sign in manually. The session is saved to `auth-state.json` and reused on future runs.

## Customising selectors

The script includes default selectors for the D365 Copilot side-panel chat. If your environment differs, update the selectors in `tests/d365-copilot-test.spec.ts` — look for comments marked `// Update these selectors`.

## Configuration (.env overrides)

The interactive prompts handle most settings, but you can also set these in `.env` to skip prompts or tune behaviour:

| Variable | Default | Description |
|---|---|---|
| `D365_URL` | *(prompted)* | Pre-fills the URL prompt on first run |
| `EDGE_PROFILE` | *(prompted)* | If set, skips the profile picker and uses this profile directly |
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
