# Repository Guidelines

## Project Structure & Module Organization
- `src/server.js` hosts the Express API, Gemini parsing logic, and QuickBooks OAuth/metadata helpers.
- `public/` contains the browser client (`index.html`, `app.js`, `styles.css`) that drives uploads, company management, and matching.
- `data/` stores runtime JSON (`parsed_invoices.json`, `quickbooks/`, `quickbooks_companies.json`); these files are generated on demand and must stay out of version control.

## Build, Test, and Development Commands
- `npm install` — install server and client dependencies (Express, Multer, node-fetch, dotenv).
- `npm start` — launch the combined API + static client on `http://localhost:3000`, warm QuickBooks metadata, and serve the UI.
- `npm test` — placeholder script; currently prints “No automated tests configured.” Treat as a reminder to add coverage.

## Coding Style & Naming Conventions
- JavaScript uses two-space indentation, semicolons, CommonJS `require`, and `async/await` for asynchronous flows.
- Prefer descriptive camelCase for variables and functions; use UPPER_SNAKE_CASE for environment-driven constants (e.g., `QUICKBOOKS_API_BASE_URL`).
- Keep helper functions near related logic in `src/server.js`; split into additional modules under `src/` only when responsibilities grow substantially.

## Testing Guidelines
- No automated suite exists yet. When adding tests, place them under `tests/`, mock external services (Gemini, QuickBooks), and wire them into `npm test`.
- Name test files after the module under test (`server.test.js`, etc.) and keep fixtures free of live customer data.
- Manual verification: run `npm start`, upload a sample invoice, confirm parsed JSON, then refresh QuickBooks metadata and validate vendor/account/tax lists.

## Commit & Pull Request Guidelines
- Use imperative commit messages (`Add QuickBooks metadata caching`, `Fix matcher state initialization`). Avoid bundling unrelated changes.
- PRs should describe the change, list manual verification steps, call out new environment variables, and attach UI screenshots or API excerpts when behavior shifts.
- Flag any security-impacting adjustments (OAuth scopes, data storage) and ensure secrets remain in `.env`, not in tracked files.

## Security & Configuration Tips
- `.env` provides Gemini and QuickBooks credentials plus overrides for `PORT` and API URLs. Never commit real keys.
- QuickBooks production access requires `QUICKBOOKS_ENVIRONMENT=production` and the corresponding base URL. Restart the server after updating environment variables.

## MCP Servers & Usage
- **Playwright MCP** (`npx @playwright/mcp@latest`) — launch when you need a live browser to validate the hosted UI (for example, reproducing issues on `https://invoice.stivescornwall.net`), capture console output, or walk through QuickBooks connect flows.
- **Context7 MCP** (`npx -y @upstash/context7-mcp --api-key …`) — fetch authoritative documentation without leaving the CLI; use it to confirm QuickBooks API semantics or third-party library details while coding.
- **Serena MCP** (`uvx --from git+https://github.com/oraios/serena serena start-mcp-server --context codex`) — start this agent when you need workflow assistance beyond simple lookups, such as synthesising multi-step reasoning with repository context and external knowledge.
- All servers are preconfigured in `/home/jay/.codex/config.toml`; activate only those relevant to your immediate task to conserve local resources.
