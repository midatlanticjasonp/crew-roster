// Storage adapters. Same interface for both:
//   init() -> {mode, user?}
//   loadMeta() / saveMeta(meta)      meta = {settings, vessels, weekIds}
//   loadCrew() / saveCrew(crewArray)
//   loadWeek(id) / saveWeek(week)    id = week ending 'YYYY-MM-DD'
import { firebaseConfig, firebaseVersion } from "./firebase-config.js";

const clean = (o) => JSON.parse(JSON.stringify(o ?? null)); // Firestore rejects undefined

class LocalStore {
  constructor() { this.mode = "local"; this.k = "crewroster."; }
  async init() { return { mode: this.mode }; }
  _get(key) { try { const v = localStorage.getItem(this.k + key); return v ? JSON.parse(v) : null; } catch { return null; } }
  _set(key, val) { localStorage.setItem(this.k + key, JSON.stringify(val)); }
  async loadMeta() { return this._get("meta"); }
  async saveMeta(m) { this._set("meta", clean(m)); }
  async loadCrew() { return this._get("crew") || []; }
  async saveCrew(c) { this._set("crew", clean(c)); }
  async loadWeek(id) { return this._get("week." + id); }
  async saveWeek(w) { this._set("week." + w.weekEnding, clean(w)); }
}

class FirebaseStore {
  constructor(cfg) { this.cfg = cfg; this.mode = "firebase"; }
  async init(onNeedSignIn) {
    const v = firebaseVersion;
    const appMod = await import(`https://www.gstatic.com/firebasejs/${v}/firebase-app.js`);
    const authMod = await import(`https://www.gstatic.com/firebasejs/${v}/firebase-auth.js`);
    const fs = await import(`https://www.gstatic.com/firebasejs/${v}/firebase-firestore.js`);
    this.fs = fs;
    const app = appMod.initializeApp(this.cfg);
    this.auth = authMod.getAuth(app);
    this.authMod = authMod;
    // Offline cache: keeps working on a bad connection and syncs when back online.
    try {
      this.db = fs.initializeFirestore(app, { localCache: fs.persistentLocalCache() });
    } catch {
      this.db = fs.getFirestore(app);
    }
    const user = await new Promise((resolve) => {
      const unsub = authMod.onAuthStateChanged(this.auth, (u) => { if (u) { unsub(); resolve(u); } else onNeedSignIn(); });
    });
    this.user = user;
    return { mode: this.mode, user: { uid: user.uid, email: user.email, name: user.displayName } };
  }
  async signIn() {
    const p = new this.authMod.GoogleAuthProvider();
    await this.authMod.signInWithPopup(this.auth, p);
  }
  async signOut() { await this.authMod.signOut(this.auth); location.reload(); }
  _doc(...path) { return this.fs.doc(this.db, ...path); }
  async _read(...path) { const s = await this.fs.getDoc(this._doc(...path)); return s.exists() ? s.data() : null; }
  async _write(data, ...path) { await this.fs.setDoc(this._doc(...path), clean(data)); }
  async loadMeta() { return this._read("app", "meta"); }
  async saveMeta(m) { await this._write(m, "app", "meta"); }
  async loadCrew() { const d = await this._read("app", "crew"); return d ? d.list : []; }
  async saveCrew(c) { await this._write({ list: c }, "app", "crew"); }
  async loadWeek(id) { return this._read("weeks", id); }
  async saveWeek(w) { await this._write(w, "weeks", w.weekEnding); }
}

export function createStore() {
  return firebaseConfig && firebaseConfig.apiKey ? new FirebaseStore(firebaseConfig) : new LocalStore();
}
