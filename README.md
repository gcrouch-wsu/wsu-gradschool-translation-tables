# WSU Gradschool Translation Tables

Export and edit **Outcomes translation tables** from pasted data. This is a **Next.js** app, split out from the former WSU Graduate School tools monorepo as its own repository.

**Repository:** [github.com/gcrouch-wsu/wsu-gradschool-translation-tables](https://github.com/gcrouch-wsu/wsu-gradschool-translation-tables)

## What it does

- Work with translation table data (including Excel-oriented flows via **xlsx** / **jszip** where used in the app).
- **Translation table validator** at `/validate`: validate an existing Outcomes-to-myWSU translation table, create a new one, or preview the join before publishing. It is a static page that runs entirely in the browser; uploaded files are not sent to the server. See [validator/USER_GUIDE.md](validator/USER_GUIDE.md).

## Requirements

- **Node.js** 18+ (20+ recommended)
- **npm**

## Quick start

```bash
git clone https://github.com/gcrouch-wsu/wsu-gradschool-translation-tables.git
cd wsu-gradschool-translation-tables
npm install
npm run dev
```

The dev script uses **port 3003** by default. Adjust `package.json` if that port is busy.

```bash
npm run build
npm run start
```

For production (e.g. Railway), use **`next start`** without a fixed `-p` so **`PORT`** from the host is used.

## Project structure

Typical layout at repo root:

```
wsu-gradschool-translation-tables/
|-- app/                 # Next.js translation tables app
|-- public/validate/     # Translation table validator (static page, served at /validate)
|-- validator/           # Validator tests, user guide, QA reports, and audit script
|-- package.json
|-- next.config.js
|-- tailwind.config.ts
|-- tsconfig.json
```

## Available scripts

| Script | Description |
|--------|-------------|
| `npm run dev` | Development server (port 3003) |
| `npm run build` | Production build |
| `npm run start` | Production server |
| `npm run lint` | ESLint |
| `npm run format` | Prettier write |
| `npm run checkfmt` | Prettier check |
| `npm run test:validator` | Validator checks, export tests, and UI smoke tests |

## Deployment

Production is the Vercel project `wsu-gradschool-translation-tables`, deployed from `main` of this repository with the repository root as the app root (https://wsu-gradschool-translation-tables.vercel.app). The validator is at `/validate` on the same site.

Any Node host that runs `npm run build` and `npm run start` with `PORT` set also works.

## Related tools

Other WSU Graduate School tools can be documented separately from this repo.

## Environment variables

None required for typical operation unless you add features that need secrets. Use `.env.local` locally and never commit secrets.
