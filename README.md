# Production Tracker — Power BI custom visual

A Power BI custom visual for shift-based production tracking and OEE
(availability × performance × quality), with a daily report view. Part of the
**Production Visual** suite (https://productionvisual.com).

## Build

```bash
npm install
npm run package      # pbiviz package -> dist/*.pbiviz
```

## Lint

```bash
npm run eslint       # npx eslint . --ext .js,.jsx,.ts,.tsx
```

## Certification notes

- **API version:** 5.10.0 (latest).
- **No external service access.** `privileges` is empty (`[]`); the visual makes
  no HTTP/S or WebSocket requests and uses no `fetch` / `XMLHttpRequest`.
- **Rendering Events API** is supported (`renderingStarted` / `renderingFinished`
  / `renderingFailed`).
- No `eval` / `Function` / `innerHTML`; the DOM is built with `createElement`.
- Source is plain TypeScript (not minified) under `src/`.
- `npm audit` reports 0 vulnerabilities; ESLint passes with no errors.
- The **`certification`** branch matches the package submitted to Partner Center.
