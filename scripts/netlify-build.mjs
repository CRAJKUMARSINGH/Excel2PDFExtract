// Resolve the __BACKEND_URL__ placeholder in web/_redirects from the
// BACKEND_URL environment variable.
//
// Netlify does not expand variables inside redirect rules, so this runs as
// the Netlify build command (see netlify.toml). Using Node instead of sed
// keeps the substitution working on Windows as well as Netlify's builder.
import { readFileSync, writeFileSync } from 'node:fs';

const FILE = new URL('../web/_redirects', import.meta.url);
const backend = process.env.BACKEND_URL;

if (!backend) {
  console.error('BACKEND_URL is not set. Add it under Netlify →');
  console.error('Site configuration → Environment variables, then redeploy.');
  process.exit(1);
}

const origin = backend.replace(/\/+$/, ''); // tolerate a trailing slash
const resolved = readFileSync(FILE, 'utf8').replaceAll('__BACKEND_URL__', origin);
writeFileSync(FILE, resolved);
console.log(`netlify-build: injected BACKEND_URL=${origin} into web/_redirects`);
