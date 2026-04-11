# Outcomes Translation Tables

Export and edit **Outcomes translation tables** from pasted data. This is a **Next.js** app, split out from the former [WSU Graduate School Tools](https://github.com/gcrouch-wsu/WSU-Mail-Editor) monorepo as its own repository.

**Repository:** [github.com/gcrouch-wsu/outcomes-translation-tables](https://github.com/gcrouch-wsu/outcomes-translation-tables)

## What it does

- Work with translation table data (including Excel-oriented flows via **xlsx** / **jszip** where used in the app).

## Requirements

- **Node.js** 18+ (20+ recommended)
- **npm**

## Quick start

```bash
git clone https://github.com/gcrouch-wsu/outcomes-translation-tables.git
cd outcomes-translation-tables
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
outcomes-translation-tables/
|-- app/
|-- components/          # if present
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

## Deployment

### Current production (legacy Vercel)

Originally deployed from the monorepo with root directory `apps/translation-tables`:

**https://wsu-mail-editor-translation-tables.vercel.app/**

### This repository

Deploy from **this** repo with **repository root** as the app root. Target: **Railway** or any Node host that runs `npm run build` and `npm run start` with `PORT` set.

## Related tools

Other WSU Graduate School tools: see [WSU-Mail-Editor](https://github.com/gcrouch-wsu/WSU-Mail-Editor).

## Environment variables

None required for typical operation unless you add features that need secrets. Use `.env.local` locally and never commit secrets.
