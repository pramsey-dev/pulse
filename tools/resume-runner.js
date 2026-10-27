/* Pulse resume runner. Paste into an employers.indeed.com tab (signed in), then click the
   "Open Pulse relay" button it adds, then call: await PR.daily(['search-id', ...]).
   Each resume goes Indeed -> relay tab (pramsey-dev.github.io/pulse/relay.html) -> Pulse server
   -> Google Drive "Candidate Resumes/<search>" folder, and its Drive link replaces the
   tracker's Resume link. Candidates that already have a Drive link are skipped. */
(function () {
  // Where Pulse is hosted. On switch day, change to the Netlify address and set PATH to ''.
  const ORIG = 'https:/' + '/pramsey-dev.github.io';
  const PATH = '/pulse';
  const norm = s => String(s || '').toLowerCase().split(',')[0].replace(/\b(dr|mr|mrs|ms|jr|sr|esq|ii|iii|phd|mba)\b\.?/g, '').replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const frame = src => { const f = document.createElement('iframe'); f.style.cssText = 'position:fixed;left:-5000px;top:0;width:1200px;height:900px'; f.src = src; document.body.appendChild(f); return f; };
  window.PR = {
    relay: null, log: [], status: 'idle', done: 0, total: 0,
    open() { this.relay = window.open(ORIG + PATH + '/relay.html?v=' + Date.now(), 'pulseRelay'); return !!this.relay; },
    call(msg, type) {
      return new Promise((res, rej) => {
        const id = Math.random().toString(36).slice(2); msg.id = id;
        const h = e => { if (e.origin !== ORIG) return; if (e.data && e.data.id === id && e.data.type === type) { window.removeEventListener('message', h); clearTimeout(t); res(e.data); } };
        const t = setTimeout(() => { window.removeEventListener('message', h); rej(new Error('relay timeout')); }, 180000);
        window.addEventListener('message', h); this.relay.postMessage(msg, ORIG);
      });
    },
    ping() {
      return new Promise(res => {
        const h = e => { if (e.origin === ORIG && e.data && e.data.type === 'pulse-pong') { window.removeEventListener('message', h); res(e.data); } };
        window.addEventListener('message', h); try { this.relay.postMessage({ type: 'pulse-ping' }, ORIG); } catch (e) {} setTimeout(() => res(null), 5000);
      });
    },
    cid(link) { const m = String(link || '').match(/(?:\?|%3F|&|%26)id(?:=|%3D)([0-9a-f]{8,})/i); return m ? m[1] : ''; },
    async resumeFor(cid) {
      const f = frame('/candidates/view?id=' + cid);
      try {
        for (let i = 0; i < 50; i++) {
          await wait(500); let a = null;
          try { a = [...f.contentDocument.querySelectorAll('a')].find(e => /Download resume/.test(e.textContent) && /^blob:/.test(e.href)); } catch (e) {}
          if (a) {
            const b = await fetch(a.href).then(r => r.blob());
            const b64 = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(String(fr.result).split(',')[1]); fr.readAsDataURL(b); });
            return { b64, filename: a.getAttribute('download') || 'resume.pdf', size: b.size };
          }
        }
        return null;
      } finally { f.remove(); }
    },
    async findId(name) {
      const n = norm(name), parts = n.split(' '), last = parts[parts.length - 1];
      const f = frame('/candidates?statusName=All&tab=manage&q=' + encodeURIComponent(last) + '&candidateName=' + encodeURIComponent(last));
      try {
        for (let i = 0; i < 40; i++) {
          await wait(500);
          let as = []; try { as = [...f.contentDocument.querySelectorAll('a[href*="candidates/view"]')]; } catch (e) {}
          if (as.length || (i > 12 && f.contentDocument && /All applications 0/.test(f.contentDocument.body.innerText))) {
            const hit = as.find(a => norm(a.textContent) === n) || as.find(a => { const t = norm(a.textContent).split(' '); return t[0] === parts[0] && t[t.length - 1] === last; });
            return hit ? (hit.href.match(/[?&]id=([0-9a-f]+)/) || [])[1] || '' : '';
          }
        }
        return '';
      } finally { f.remove(); }
    },
    async copy(searchId, r, cid, how) {
      let x = null; try { x = await this.resumeFor(cid); } catch (e) {}
      if (!x) { this.log.push([searchId, r.name, 'no resume on Indeed']); return; }
      try {
        const s = await this.call({ type: 'pulse-resume', payload: { searchId, name: r.name, filename: x.filename, b64: x.b64, indeedUrl: 'https:/' + '/employers.indeed.com/candidates/view?id=' + cid } }, 'pulse-resume-result');
        this.log.push([searchId, r.name, s.ok ? 'saved row ' + s.result.row + (how ? ' (' + how + ')' : '') : 'error ' + s.error]);
      } catch (e) { this.log.push([searchId, r.name, 'error ' + e.message]); }
    },
    async daily(searchIds) {
      this.log = []; this.done = 0; this.total = 0;
      for (const searchId of searchIds) {
        this.status = 'loading ' + searchId;
        const t = await this.call({ type: 'pulse-targets', searchId }, 'pulse-targets-result');
        if (!t.ok) { this.log.push([searchId, '', 'could not read tracker: ' + t.error]); continue; }
        const rows = t.result.rows.filter(r => !r.hasCopy); this.total += rows.length;
        for (const r of rows) {
          this.status = searchId + ': ' + r.name;
          let cid = this.cid(r.link), how = '';
          if (!cid) { cid = await this.findId(r.name); how = 'found by name'; }
          if (!cid) this.log.push([searchId, r.name, 'not found on Indeed']); else await this.copy(searchId, r, cid, how);
          this.done++;
        }
      }
      this.status = 'finished';
      return this.log;
    }
  };
  const b = document.createElement('button'); b.id = 'prOpen'; b.textContent = 'Open Pulse relay';
  b.style.cssText = 'position:fixed;right:20px;bottom:20px;z-index:99999;padding:14px 18px;font:600 16px system-ui;background:#007f91;color:#fff;border:0;border-radius:10px';
  b.onclick = () => { PR.open(); b.textContent = 'Relay opened'; }; document.body.appendChild(b);
})();
