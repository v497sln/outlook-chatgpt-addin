// Bundles the app into one self-contained HTML file (dist/squadsage.html):
// inline CSS, inline JS, and the solver worker embedded as a Blob source.
// Usage: node build.mjs
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';

const read = f => readFileSync(new URL(f, import.meta.url), 'utf8');
const ORDER = ['data', 'engine', 'store', 'app'];

// Turn an ES module into an IIFE that registers its exports in __m[name].
function wrap(name, code) {
  const exportsList = [];
  code = code.replace(/^import \* as (\w+) from '\.\/(\w+)\.js';$/gm, (_, n, mod) => `const ${n} = __m.${mod};`);
  code = code.replace(/^import \{([^}]+)\} from '\.\/(\w+)\.js';$/gm, (_, names, mod) => `const {${names}} = __m.${mod};`);
  code = code.replace(/^export (async function|function|const|let) (\w+)/gm, (_, kw, n) => { exportsList.push(n); return `${kw} ${n}`; });
  if (/^\s*(import|export)\s/m.test(code)) throw new Error(`Unhandled import/export in ${name}`);
  return `__m.${name} = (() => {\n${code}\nreturn { ${exportsList.join(', ')} };\n})();\n`;
}

const modules = Object.fromEntries(ORDER.map(n => [n, read(`js/${n}.js`)]));

const workerSrc = `const __m = {};\n${wrap('engine', modules.engine)}
self.onmessage = (e) => {
  try { self.postMessage({ ok: true, result: __m.engine.solve(e.data) }); }
  catch (err) { self.postMessage({ ok: false, error: String(err && err.message || err) }); }
};`;

const appJs = `globalThis.__SOLVER_SRC = ${JSON.stringify(workerSrc).replace(/<\//g, '<\\/')};
const __m = {};
${ORDER.map(n => wrap(n, modules[n])).join('\n')}`;

const html = read('index.html');
const body = html.slice(html.indexOf('<body>') + 6, html.indexOf('<script type="module"'));
const css = read('styles.css');
const icon = `data:image/svg+xml,${encodeURIComponent(read('icon.svg').trim())}`;

const out = `<title>SquadSage</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>
${css}
.top { top: env(safe-area-inset-top, 0px); padding-top: 10px; }
body::before { content: ""; position: fixed; inset: 0 0 auto; height: env(safe-area-inset-top, 0px); z-index: 10; background: var(--bg); }
body { background: var(--bg); }
</style>
${body.replace(/src="icon\.svg"/g, `src="${icon}"`).trim()}
<script type="module">
${appJs.replace(/<\/script/gi, '<\\/script')}
</script>
`;

mkdirSync(new URL('dist/', import.meta.url), { recursive: true });
writeFileSync(new URL('dist/squadsage.html', import.meta.url), out);
console.log(`dist/squadsage.html  ${(out.length / 1024).toFixed(1)} KB`);
