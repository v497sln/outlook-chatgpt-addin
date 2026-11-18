// Copies the web app into www/ for Capacitor (the Android app's web layer).
import { cpSync, rmSync, mkdirSync } from 'node:fs';
const root = new URL('./', import.meta.url);
const www = new URL('www/', root);
rmSync(www, { recursive: true, force: true });
mkdirSync(www, { recursive: true });
for (const f of ['index.html', 'styles.css', 'icon.svg', 'manifest.webmanifest', 'js']) {
  cpSync(new URL(f, root), new URL(f, www), { recursive: true });
}
console.log('www/ ready');
