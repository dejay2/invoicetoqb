# Repository Guidelines

## Project Structure & Module Organization
- `src/server.js` hosts the Express API, Gemini integration, and QuickBooks OAuth helpers.
- `public/` serves the static web client (`index.html`, `app.js`, `styles.css`) that uploads invoices and manages QuickBooks connections.
- `data/parsed_invoices.json` and `data/quickbooks_companies.json` persist parsed invoice history and OAuth tokens; these files are created on demand and should be git-ignored.
- Use a `.env` file in the project root to supply Gemini and QuickBooks credentials plus any overrides for `PORT` or API base URLs.

## Build, Test, and Development Commands
- `npm install` resolves runtime dependencies (`express`, `multer`, `node-fetch`, `dotenv`).
- `npm start` launches the server on `http://localhost:3000` and serves both the API and static client.
- `npm test` currently reports "No automated tests configured"; treat this as a reminder to add coverage when extending the project.
- During development, watch server logs for parsing errors and duplicate detection feedback.

## Configuration & Security Notes
- Required environment variables: `GEMINI_API_KEY`, `QUICKBOOKS_CLIENT_ID`, `QUICKBOOKS_CLIENT_SECRET`; optional overrides include `GEMINI_MODEL`, `QUICKBOOKS_*` endpoints, and `PORT`.
- Keep OAuth tokens and parsed invoice data in the `data/` directory; never commit real customer data. Rotate API keys regularly and clear obsolete `quickBooksStates` by restarting the server.

## Coding Style & Naming Conventions
- Follow the existing Node.js style: two-space indentation, semicolons, CommonJS `require`, and `async/await` for async work.
- Prefer descriptive camelCase for variables and functions; use uppercase snake case for environment-driven constants.
- Encapsulate new logic in small helpers alongside related functions inside `src/server.js`, or create new modules under `src/` when responsibilities grow.

## Testing Guidelines
- Manual verification: run `npm start`, upload a sample invoice through the UI, and confirm JSON output plus duplicate detection in the UI.
- Use `curl -F "invoice=@sample.pdf" http://localhost:3000/api/parse-invoice` for quick API checks.
- When adding automated tests, colocate them under a new `tests/` directory and mock external services (Gemini, QuickBooks) to keep runs deterministic.

## Commit & Pull Request Guidelines
- Commit messages should use an imperative summary (`Add QuickBooks token refresh`) followed by optional detail in the body.
- Ensure PRs describe the change, list manual verification steps, and flag any new environment variables or schema changes.
- Attach screenshots or terminal excerpts when UI flow or API responses change, and request review on security-impacting updates.
