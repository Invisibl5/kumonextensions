# kumonextensions

Tampermonkey script: **Kumon Auto Grader** (mark failure boxes with X/Triangle, clear, reapply from results). Auto-updates from this repo.

## Install (one-time)

1. Open in browser:  
   **https://raw.githubusercontent.com/Invisibl5/kumonextensions/main/kumonextensions.user.js**
2. Tampermonkey will prompt → click **Install**.
3. Reload Kumon Connect (e.g. `https://class-navi.digital.kumon.com/...`). The panel appears top-right.

## Update workflow

- **Edit** `kumonextensions.user.js` in Cursor.
- **Bump** `@version` in the script header when you want users to get the update (e.g. `0.2.1` → `0.2.2`).
- **Commit & push** to `main`. Tampermonkey will auto-update installed copies (by checking `@updateURL`).

## Run only on Kumon (optional)

The script is set to run only on:

- `https://class-navi.digital.kumon.com/us/index.html`

To run on all Kumon Class-Navi US pages, change the script header to:

- `// @match        https://class-navi.digital.kumon.com/us/*`
