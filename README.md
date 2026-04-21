# D365 Customer Service Copilot Test

Playwright-based test that sends each prompt from `Prompts and Responses.xlsx` to the D365 Customer Service Copilot, captures the response, compares it against the expected answer, and writes results to a new Excel file.

## Setup

1. **Configure the D365 URL** — edit `.env` and set `D365_URL`:
   ```
   D365_URL=https://yourorg.crm.dynamics.com
   ```

2. **Install dependencies** (already done if you cloned this):
   ```
   npm install
   ```

3. **Authenticate** — choose one of two methods:

   **Option A – Edge profile (recommended):**
   Set `EDGE_PROFILE` in `.env` to a profile that is already signed in to D365.
   To find the name, open `edge://version` in Edge and look at **Profile path** — use the last folder name (e.g. `Default`, `Profile 1`).
   ```
   EDGE_PROFILE=Profile 1
   ```

   **Option B – Manual login (fallback):**
   Leave `EDGE_PROFILE` blank and run:
   ```
   npx ts-node auth-setup.ts
   ```
   Sign in to D365 in the browser that opens, then press Enter in the terminal.

4. **Adjust selectors** — the script has default selectors for the D365 Copilot chat panel. If your environment uses different selectors, update them in `tests/d365-copilot-test.spec.ts` (search for `// Update these selectors`).

## Running the Test

```
npx playwright test
```

This runs in **headed** mode (visible browser) by default. Results are saved to `Test Results YYYY-MM-DD.xlsx`.

## Configuration (.env)

| Variable | Default | Description |
|---|---|---|
| `D365_URL` | *(must set)* | Your D365 Customer Service URL |
| `EDGE_PROFILE` | *(empty)* | Edge profile name (e.g. `Default`, `Profile 1`). If set, uses Edge with that profile's cookies/session. If blank, falls back to `auth-state.json`. |
| `COPILOT_RESPONSE_TIMEOUT` | `60` | Seconds to wait for each Copilot response |
| `SIMILARITY_THRESHOLD` | `0.6` | 0–1 word-overlap threshold for pass/fail |

## Output

The output Excel contains:
- **#** — prompt number
- **Prompt** — the question sent
- **Expected Response** — from the source spreadsheet
- **Actual Response** — what Copilot returned
- **Similarity** — word-overlap percentage
- **Result** — PASS (green) or FAIL (red)
- **Referenced Docs** — from the source spreadsheet
