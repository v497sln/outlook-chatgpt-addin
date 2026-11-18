// Builds www/ for the Android app (Capacitor):
//   www/index.html        <- shell.html (opens the EA Web App full-screen)
//   www/overlay.bundle.js <- engine.js + overlay.js, injected into the EA page
import { cpSync, rmSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
const root = new URL('./', import.meta.url);
const www = new URL('www/', root);
const read = f => readFileSync(new URL(f, root), 'utf8');

function wrap(name, code) {
  const names = [];
  code = code.replace(/^export (async function|function|const|let) (\w+)/gm, (_, kw, n) => { names.push(n); return `${kw} ${n}`; });
  if (/^\s*(import|export)\s/m.test(code)) throw new Error(`Unhandled import/export in ${name}`);
  return `__m.${name} = (() => {\n${code}\nreturn { ${names.join(', ')} };\n})();\n`;
}

rmSync(www, { recursive: true, force: true });
mkdirSync(www, { recursive: true });
cpSync(new URL('shell.html', root), new URL('index.html', www));
cpSync(new URL('icon.svg', root), new URL('icon.svg', www));
const bundle = `(function () {\nif (window.__ssOverlay) return;\nconst __m = {};\n${wrap('engine', read('js/engine.js'))}\n${read('js/overlay.js')}\n})();\n`;
writeFileSync(new URL('overlay.bundle.js', www), bundle);
console.log(`www/ ready (overlay ${(bundle.length / 1024).toFixed(1)} KB)`);
