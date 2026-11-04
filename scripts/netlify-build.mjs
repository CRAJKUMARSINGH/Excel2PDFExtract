// Resolve the __BACKEND_URL__ placeholder in web/_redirects from the
// BACKEND_URL environment variable.
//
// Netlify does not expand variables inside redirect rules, so this runs as
// the Netlify build command (see netlify.toml). Using Node instead of sed
// keeps the substitution working on Windows as well as Netlify's builder.
//
// If BACKEND_URL is not configured the static UI is still published — the
// site must never be left 404ing — and the /api + /download proxy rules are
// dropped until a backend host is provided.
import { readFileSync, writeFileSync } from 'node:fs';

const FILE = new URL('../web/_redirects', import.meta.url);
const backend = process.env.BACKEND_URL;

let text = readFileSync(FILE, 'utf8');

if (backend) {
  const origin = backend.replace(/\/+$/, ''); // tolerate a trailing slash
  text = text.replaceAll('__BACKEND_URL__', origin);
  console.log(`netlify-build: injected BACKEND_URL=${origin} into web/_redirects`);
} else {
  console.warn('netlify-build: BACKEND_URL is not set — publishing the static UI only.');
  console.warn('netlify-build: add it under Netlify → Site configuration →');
  console.warn('netlify-build: Environment variables, then redeploy to enable /api and /download.');
  text = text
    .replace(/^\/api\/\*.*\r?\n?/m, '')
    .replace(/^\/download\/\*.*\r?\n?/m, '');
}

writeFileSync(FILE, text);
