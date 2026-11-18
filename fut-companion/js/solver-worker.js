// Runs the solver off the main thread so the UI never freezes while solving.
import { solve } from './engine.js';

self.onmessage = (e) => {
  try {
    self.postMessage({ ok: true, result: solve(e.data) });
  } catch (err) {
    self.postMessage({ ok: false, error: String(err && err.message || err) });
  }
};
