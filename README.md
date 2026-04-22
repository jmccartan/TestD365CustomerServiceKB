# D365 Customer Service Copilot KB Test Harness

An automated regression-testing tool for **Dynamics 365 Customer Service Copilot** knowledge bases. It lets you define a set of "golden" prompts with expected answers, then runs them against your live Copilot instance and reports how well each response matches.

**Why this exists:** When you add, update, or reorganise articles in a D365 Customer Service knowledge base, there's no built-in way to verify that Copilot still answers key questions correctly. This tool fills that gap — you author a spreadsheet of prompts and expected responses, run the test, and get a colour-coded Excel report showing pass/fail for every prompt, the actual Copilot response, and which KB articles were cited.

### How it works

1. Opens a Chromium browser and signs in to your D365 Customer Service workspace.
2. Reads prompts from `Prompts and Responses.xlsx`.
3. Sends each prompt to the Copilot side-panel, waits for the response to finish streaming, then expands "Check sources" to capture which KB articles were cited.
4. Compares the actual response to the expected answer using word-overlap similarity.
5. Writes a timestamped Excel report with pass/fail results, similarity scores, and cited sources.

The tool automatically handles D365's 15-message conversation limit by clearing the chat and continuing.

## Prerequisites

- **Node.js** 18+ (LTS recommended)

### Installing Node.js

If you don't have Node.js installed, pick one of these options:

| Method | Command / Link |
|---|---|
| **Windows installer** (easiest) | Download from [nodejs.org](https://nodejs.org/) — choose the LTS version and run the `.msi` installer. |
| **winget** (Windows) | `winget install OpenJS.NodeJS.LTS` |
| **Homebrew** (macOS) | `brew install node@20` |
| **nvm** (any OS) | Install [nvm](https://github.com/nvm-sh/nvm), then `nvm install --lts` |

After installing, verify with:
```
node --version   # should show v18+ (e.g. v20.x or v24.x)
npm --version    # should show 9+
```

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

## Creating Your Golden Prompts

The test is driven by `Prompts and Responses.xlsx`. This is your "golden" test suite — the questions you expect Copilot to answer correctly based on your knowledge base content.

**Spreadsheet format** (first sheet, or a sheet named "Prompts & Responses"):

| Column A — Prompt | Column B — Expected Response | Column C — Referenced Docs (optional) |
|---|---|---|
| What is our return policy? | Customers may return items within 30 days of purchase for a full refund... | Return Policy KB Article |
| How do I reset my password? | Navigate to Settings > Security > Change Password... | Account Management Guide |

**Tips for writing effective golden prompts:**

- **Cover your critical KB topics** — include at least one prompt per major article or topic area.
- **Use natural language** — write prompts the way a real agent or customer would ask them, not keyword searches.
- **Expected responses don't need to be exact** — the tool uses word-overlap similarity (default 60% threshold), so paraphrases will still pass. Focus on including the key terms and facts.
- **Include edge cases** — add prompts for topics you know are tricky, recently changed, or frequently confused.
- **Version your spreadsheet** — as your KB evolves, update the golden prompts to match. This is your regression safety net.
- **Column C is optional** — use it to note which KB article(s) should be the source, so you can cross-reference against the "Cited Sources" column in the results.

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

Results are saved to `Test Results YYYY-MM-DD_HH-MM-SS.xlsx` with these columns:

| Column | Description |
|---|---|
| **#** | Prompt number |
| **Prompt** | The question sent to Copilot |
| **Expected Response** | From the source spreadsheet |
| **Actual Response** | What Copilot returned |
| **Similarity** | Word-overlap percentage |
| **Result** | PASS (green) or FAIL (red) |
| **Referenced Docs** | From the source spreadsheet |
| **Cited Sources** | KB articles cited by Copilot (from "Check sources") |

A summary row at the bottom shows total pass/fail counts.
