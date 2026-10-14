/* Vonnue Cup site. Loads the organiser's workbook, then renders every view from it. */
(function () {
  'use strict';

  const USER = window.VONNUE_CONFIG || {};
  const CONFIG = {
    brackets: 'data/Vonnue_Cup_Knockout_Brackets.xlsx',
    participants: 'data/Vonnue_Cup_Participants.xlsx',
    snapshot: 'data/snapshot.json',
    googleSheet: USER.googleSheet || '',
    googleParticipantsSheet: USER.googleParticipantsSheet || '',
    year: USER.year || 2026,
    refreshMs: (USER.refreshMinutes || 2) * 60 * 1000,
  };
  const refreshLabel = `${CONFIG.refreshMs / 60000} minute${CONFIG.refreshMs === 60000 ? '' : 's'}`;

  // Turn a Google Sheets link (edit link or "Publish to web" link) into its .xlsx download link.
  function googleXlsxUrl(link) {
    if (!link) return null;
    const pub = link.match(/spreadsheets\/d\/e\/([\w-]+)/);
    if (pub) return `https://docs.google.com/spreadsheets/d/e/${pub[1]}/pub?output=xlsx`;
    const id = link.match(/spreadsheets\/d\/([\w-]+)/);
    return id ? `https://docs.google.com/spreadsheets/d/${id[1]}/export?format=xlsx` : null;
  }

  let M = null;            // parsed model
  let sourceLabel = '';
  let lastSig = '';
  const state = { sport: 'All', team: 'all', hideDone: false, q: '' };

  const $ = (s, el = document) => el.querySelector(s);
  const view = $('#view');

  // ---------- helpers
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const slug = (s) => String(s).toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  const now = () => {
    const q = new URLSearchParams(location.search).get('now');
    const d = q ? new Date(q) : null;
    return d && !isNaN(d) ? d : new Date();
  };
  const DAY = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const keyOf = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const dateOfKey = (k) => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d); };
  const shortDate = (d) => `${DAY[d.getDay()].slice(0, 3)} ${d.getDate()} ${MON[d.getMonth()].slice(0, 3)}`;
  const plural = (n, w, p) => `${n} ${n === 1 ? w : (p || w + 's')}`;

  const ICONS = {
    'Chess': '<circle cx="12" cy="6.3" r="2.4"/><path d="M9.6 10.2h4.8M10.4 10.2c0 3.1-1.3 5-2.4 6.6h8c-1.1-1.6-2.4-3.5-2.4-6.6M6.5 20.2h11v-3.4h-11z"/>',
    'Carrom': '<rect x="3.5" y="3.5" width="17" height="17" rx="2.5"/><circle cx="12" cy="12" r="3.2"/><circle cx="7" cy="7" r="1.1"/><circle cx="17" cy="7" r="1.1"/><circle cx="7" cy="17" r="1.1"/><circle cx="17" cy="17" r="1.1"/>',
    'Table Tennis': '<circle cx="10" cy="10" r="6.2"/><path d="M14.4 14.4l5.6 5.6"/><circle cx="18.6" cy="5.4" r="1.7"/>',
    'Badminton': '<circle cx="6.6" cy="17.4" r="2.4"/><path d="M8.3 15.7L19.5 4.5M8.3 15.7l9.4-6.3 1.8-4.9-4.9 1.8z M12 12l2.6-5.1M12 12l5.1-2.6"/>',
    'Cricket': '<path d="M15.8 3.2l5 5-9.9 9.9-5-5z"/><path d="M5.9 13.1L3 16l5 5 2.9-2.9"/><circle cx="6.2" cy="6.2" r="2.3"/>',
    'Football': '<circle cx="12" cy="12" r="8.6"/><path d="M12 7.6l3.4 2.5-1.3 4h-4.2l-1.3-4z M12 7.6V3.4 M15.4 10.1l4-1.3 M14.1 14.1l2.5 3.4 M9.9 14.1l-2.5 3.4 M8.6 10.1l-4-1.3"/>',
    trophy: '<path d="M7 4h10v4a5 5 0 0 1-10 0z M7 6H4.5a2.5 2.5 0 0 0 2.6 3.6 M17 6h2.5a2.5 2.5 0 0 1-2.6 3.6 M12 13v3.5 M8.5 20h7l-.8-3.5h-5.4z"/>',
    calendar: '<rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 10h17M8 3v4M16 3v4"/>',
    users: '<circle cx="9" cy="8.5" r="3.3"/><path d="M3 19.5c.6-3.3 3-5 6-5s5.4 1.7 6 5M15.5 5.6a3.2 3.2 0 0 1 0 6M17.5 14.6c1.9.6 3.1 2.2 3.5 4.9"/>',
    pin: '<path d="M12 21s-6.5-6.1-6.5-11a6.5 6.5 0 0 1 13 0c0 4.9-6.5 11-6.5 11z"/><circle cx="12" cy="10" r="2.3"/>',
    clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
  };
  const icon = (k) => `<svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[k] || ICONS.trophy}</svg>`;
  const SPORTS = ['Chess', 'Carrom', 'Table Tennis', 'Badminton', 'Cricket', 'Football'];

  const team = (id) => (M && M.teamById[id]) || null;
  const tc = (t) => (t ? `--tc: var(--t${t.id})` : '--tc: var(--muted)');
  const initials = (name) => {
    const w = String(name).replace(/[^A-Za-z ]/g, ' ').trim().split(/\s+/);
    return (w.length > 1 ? w[0][0] + w[1][0] : w[0].slice(0, 2)).toUpperCase();
  };
  // Team logos live in assets/teams; a crest falls back to the team's initials if an image fails.
  const logoSrc = (t, big) => `assets/teams/team-${t.id}${big ? '' : '-sm'}.webp`;
  const crest = (t, size, big) => t ? `<span class="crest has-logo" style="${tc(t)}${size ? `;--size:${size}px` : ''}" title="${esc(t.name)}"><img src="${logoSrc(t, big || size >= 64)}" alt="" loading="lazy" decoding="async"><span class="ini">${initials(t.name)}</span></span>` : '';
  const teamTag = (t) => t ? `<span class="team-tag" style="${tc(t)}"><i class="tdot"></i><b>${esc(t.name)}</b></span>` : '';

  // ---------- match status
  function status(m) {
    if (m.winner || m.isDraw) return 'done';
    const t = now();
    const timed = m.start && /^\d/.test(m.start);
    if (timed && m.startAt && m.endAt && t >= m.startAt && t <= m.endAt) return 'live';
    const dayEnd = m.startAt ? new Date(m.startAt.getFullYear(), m.startAt.getMonth(), m.startAt.getDate(), 23, 59) : null;
    const over = timed && m.endAt ? t > m.endAt : dayEnd && t > dayEnd;
    if (over) return 'pending';
    return 'upcoming';
  }
  const statusPill = (s) => ({
    live: '<span class="pill live">Live</span>',
    done: '<span class="pill done">Final</span>',
    pending: '<span class="pill pending">Result due</span>',
    upcoming: '<span class="pill">Upcoming</span>',
  }[s]);

  function slotLabel(s) {
    if (s.name) return s.name;
    const src = s.src;
    if (src.kind === 'winner') return `Winner ${src.match}`;
    if (src.kind === 'loser') return `Loser ${src.match}`;
    if (src.kind === 'pool') return s.label && /^Tie/.test(s.label) ? `${src.label} (tie)` : src.label;
    return s.label || 'TBD';
  }

  function eventOf(m) { return M.eventByName[m.event] || null; }
  function sportOfMatch(m) { const e = eventOf(m); return e ? e.sport : 'Other'; }

  // ---------- match row
  function matchRow(m, opts = {}) {
    const st = status(m);
    const ev = eventOf(m);
    const side = (s, right) => {
      const t = M.teamOf(m.event, s.name);
      const cls = !s.name ? 'tbd' : m.winner ? (m.winner === s.name ? 'win' : 'lose') : '';
      return `<div class="side ${right ? 'right' : ''} ${cls}" style="${tc(t)}"><i class="tdot"></i><span title="${esc(slotLabel(s))}${t ? ' · ' + esc(t.name) : ''}">${esc(slotLabel(s))}</span></div>`;
    };
    const timeTop = m.start && /^\d/.test(m.start) ? m.start : (m.start ? 'Then' : '—');
    const timeSub = opts.showDate ? shortDate(m.startAt) : (m.end ? `to ${m.end}` : (m.start === 'follows' ? 'follows' : ''));
    const isFinal = m.round === 'Final';
    return `<div class="mrow ${st === 'live' ? 'is-live' : ''} ${isFinal ? 'is-final' : ''}" data-mid="${esc(m.id)}">
      <div class="mtime"><b>${esc(timeTop)}</b><small>${esc(timeSub)}</small></div>
      <div class="mbody">
        <div class="mmeta">
          <a class="evname" href="#event-${ev ? ev.slug : ''}">${icon(sportOfMatch(m))}${esc(m.event)}</a>
          <span>${esc(m.round)}${m.pool ? '' : ''}</span>
          <span class="mid">${esc(m.id)}</span>
        </div>
        <div class="vs">${side(m.slots[0])}${m.score || m.isDraw ? `<span class="score">${esc([m.score, m.isDraw ? 'Draw' : ''].filter(Boolean).join(' · '))}</span>` : '<span class="sep">VS</span>'}${side(m.slots[1], true)}</div>
      </div>
      <div class="mend">${isFinal && st !== 'done' ? '<span class="pill gold">Final</span>' : statusPill(st)}<small>${icon('pin')} ${esc(m.venue || '')}</small></div>
    </div>`;
  }

  // ---------- data loading
  async function fetchBuf(url) {
    const r = await fetch(url, { cache: 'no-store' });
    if (!r.ok) throw new Error(`${url} returned ${r.status}`);
    return { buf: await r.arrayBuffer(), modified: r.headers.get('last-modified') };
  }
  function signature(buf) {
    const a = new Uint8Array(buf); let h = 2166136261;
    for (let i = 0; i < a.length; i += 7) { h ^= a[i]; h = Math.imul(h, 16777619); }
    return `${a.length}:${h >>> 0}`;
  }

  // Sources in order of preference: Google Sheets (if set), the workbook file, the JSON snapshot.
  async function loadSource(kind) {
    if (kind === 'google') {
      const gB = googleXlsxUrl(CONFIG.googleSheet), gP = googleXlsxUrl(CONFIG.googleParticipantsSheet);
      const [b, p] = await Promise.all([fetchBuf(gB), (gP ? fetchBuf(gP) : fetchBuf(CONFIG.participants)).catch(() => null)]);
      return { buf: b.buf, parts: p && p.buf, sig: signature(b.buf) + (p ? signature(p.buf) : ''), label: `Results update automatically every ${refreshLabel}` };
    }
    if (kind === 'file') {
      const [b, p] = await Promise.all([fetchBuf(CONFIG.brackets), fetchBuf(CONFIG.participants).catch(() => null)]);
      const mod = b.modified ? new Date(b.modified) : null;
      const when = mod && !isNaN(mod) ? mod.toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' }) : '';
      return { buf: b.buf, parts: p && p.buf, sig: signature(b.buf) + (p ? signature(p.buf) : ''), label: `Workbook ${when ? 'saved ' + when : 'loaded'} · checked for updates every ${refreshLabel}` };
    }
    const r = await fetch(CONFIG.snapshot, { cache: 'no-store' });
    if (!r.ok) throw new Error(`${CONFIG.snapshot} returned ${r.status}`);
    const text = await r.text();
    const snap = JSON.parse(text);
    const when = new Date(snap.savedAt).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
    return { buf: snap.brackets, parts: snap.participants, sig: signature(new TextEncoder().encode(text).buffer), label: `Snapshot of the workbook from ${when}` };
  }

  async function loadRemote(initial) {
    const kinds = [...(googleXlsxUrl(CONFIG.googleSheet) ? ['google'] : []), 'file', 'snapshot'];
    let lastErr = null;
    for (const kind of kinds) {
      try {
        const src = await loadSource(kind);
        if (src.sig === lastSig) return;
        const model = VonnueParser.parse(XLSX, src.buf, src.parts, { year: CONFIG.year });
        // A sheet that isn't the Vonnue Cup workbook (or has lost its layout) must not replace good data
        if (!model.events.length || !model.matches.length || !model.teams.length) throw new Error('This workbook has no Vonnue Cup events, matches or teams');
        lastSig = src.sig;
        M = model;
        sourceLabel = src.label;
        setStatus();
        render(!initial);
        return;
      } catch (e) {
        lastErr = e;
        if (kind === 'google') console.warn('Google Sheets could not be used, falling back to the workbook file:', e);
      }
    }
    console.error(lastErr);
    if (initial && !M) showLoadError(lastErr);
  }
  function setStatus() {
    $('#data-status').textContent = `${sourceLabel}. Last checked ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}.`;
  }

  function showLoadError(e) {
    view.innerHTML = `<div class="wrap"><div class="panel error-box">
      <span class="eyebrow">Couldn't load the tournament data</span>
      <h2>Workbook not found</h2>
      <p>The site reads <code>${esc(CONFIG.brackets)}</code> from the same folder. If you opened <code>index.html</code> straight from disk, the browser blocks that. Serve the folder instead, for example with <code>python3 -m http.server</code>, or pick the workbook below to view it now.</p>
      <p><label class="file-pick" for="file-input-2"><input id="file-input-2" type="file" accept=".xlsx" multiple> Choose Vonnue_Cup_Knockout_Brackets.xlsx</label></p>
      <p style="color:var(--muted);font-size:.8rem">Details: ${esc(e.message)}</p>
    </div></div>`;
    $('#file-input-2').addEventListener('change', onFilePick);
  }

  async function onFilePick(ev) {
    const files = [...ev.target.files];
    if (!files.length) return;
    let bracket = files.find((f) => /bracket/i.test(f.name)) || files.find((f) => !/participant/i.test(f.name)) || files[0];
    const part = files.find((f) => /participant/i.test(f.name));
    try {
      const pb = part ? await part.arrayBuffer() : null;
      M = VonnueParser.parse(XLSX, await bracket.arrayBuffer(), pb, { year: CONFIG.year });
      sourceLabel = `Showing ${bracket.name} from your device (not shared with anyone)`;
      lastSig = 'local';
      clearInterval(refreshTimer);
      $('#data-status').textContent = sourceLabel + '.';
      render();
    } catch (e) {
      $('#data-status').textContent = `${bracket.name} could not be read as the Vonnue Cup brackets workbook. Pick Vonnue_Cup_Knockout_Brackets.xlsx.`;
    }
  }

  // ---------- routing
  function route() {
    const h = (location.hash || '#home').slice(1);
    if (h.startsWith('event-')) return { name: 'brackets', ev: h.slice(6) };
    if (h.startsWith('team-')) return { name: 'team', id: +h.slice(5) };
    if (h.startsWith('player-')) return { name: 'player', key: h.slice(7) };
    if (h.startsWith('day-')) return { name: 'schedule', day: h.slice(4) };
    return { name: ['home', 'schedule', 'brackets', 'standings', 'teams'].includes(h) ? h : 'home' };
  }

  function render(keepScroll) {
    if (!M) return;
    const r = route();
    const y = window.scrollY;
    const navKey = r.name === 'team' || r.name === 'player' ? 'teams' : r.name;
    document.querySelectorAll('[data-nav]').forEach((a) => a.classList.toggle('active', a.dataset.nav === navKey));
    const html = ({ home: viewHome, schedule: viewSchedule, brackets: viewBrackets, standings: viewStandings, teams: viewTeams, team: viewTeam, player: viewPlayer }[r.name])(r);
    view.classList.toggle('no-anim', !!keepScroll);
    view.innerHTML = `<div class="wrap">${html}</div>`;
    afterRender(r);
    if (keepScroll) window.scrollTo(0, y);
    else if (r.day) { const el = document.getElementById('d-' + r.day); if (el) el.scrollIntoView(); }
    else window.scrollTo(0, 0);
  }

  // ---------- derived stats
  function nextMatches(n) {
    const t = now();
    const open = M.matches.filter((m) => !m.winner && !m.isDraw && m.startAt);
    const live = open.filter((m) => status(m) === 'live');
    const upcoming = open.filter((m) => {
      if (status(m) !== 'upcoming') return false;
      const timed = m.start && /^\d/.test(m.start);
      return timed ? m.startAt >= t : keyOf(m.startAt) >= keyOf(t);
    });
    return { live, upcoming: upcoming.slice(0, n) };
  }
  function tournamentSpan() {
    const ds = M.matches.map((m) => m.startAt).filter(Boolean).sort((a, b) => a - b);
    return { first: ds[0], last: ds[ds.length - 1] };
  }
  function ring(frac, color) {
    const r = 21, c = 2 * Math.PI * r;
    return `<svg class="stat-ring" viewBox="0 0 52 52" aria-hidden="true"><circle cx="26" cy="26" r="${r}" stroke="var(--surface-3)"/><circle cx="26" cy="26" r="${r}" stroke="${color}" stroke-dasharray="${(c * Math.min(1, frac)).toFixed(1)} ${c.toFixed(1)}" transform="rotate(-90 26 26)"/></svg>`;
  }

  // ---------- views
  function viewHome() {
    const span = tournamentSpan();
    const played = M.matches.filter((m) => m.winner || m.isDraw).length;
    const decided = M.events.filter((e) => e.champion).length;
    const { live, upcoming } = nextMatches(5);
    const nextOne = live[0] || upcoming[0];
    const playersN = M.players.length;
    const t = now();
    const daysLeft = span.last ? Math.max(0, Math.ceil((dateOfKey(keyOf(span.last)) - dateOfKey(keyOf(t))) / 864e5)) : 0;
    const started = span.first && t >= span.first;

    const countdown = nextOne && !live.length ? `<div class="countdown" id="countdown" data-at="${nextOne.startAt.getTime()}">
        <div><b data-u="d">0</b><small>Days</small></div><div><b data-u="h">0</b><small>Hours</small></div>
        <div><b data-u="m">0</b><small>Min</small></div><div><b data-u="s">0</b><small>Sec</small></div></div>` : '';
    const nextHead = live.length ? `<span class="pill live">${plural(live.length, 'match', 'matches')} live</span>`
      : nextOne ? `<span class="eyebrow">${started ? 'Next match' : 'First match'} · ${shortDate(nextOne.startAt)}${nextOne.start && /^\d/.test(nextOne.start) ? ' · ' + nextOne.start : ''}</span>`
      : '<span class="pill gold">Tournament complete</span>';
    const nextList = (live.length ? live.slice(0, 3) : upcoming.slice(0, 3)).map((m) => matchRow(m)).join('');

    const ranked = [...M.teams].sort((a, b) => (a.rank || 9) - (b.rank || 9) || b.points - a.points || a.id - b.id);
    const maxPts = Math.max(1, ...ranked.map((x) => x.points));
    const anyPts = ranked.some((x) => x.points > 0);

    return `
    <section class="hero rise">
      <div>
        <span class="eyebrow">${span.first ? `${shortDate(span.first)} – ${shortDate(span.last)} ${CONFIG.year}` : ''}</span>
        <h1 class="hero-title" style="margin-top:10px">Vonnue<span>Cup ${CONFIG.year}</span></h1>
        <p class="hero-sub">Six teams, twelve events, one trophy. Follow every match, bracket and point as the results come in.</p>
        <div class="hero-meta">
          ${SPORTS.map((s) => `<span class="chip">${icon(s)}${s}</span>`).join('')}
        </div>
      </div>
      <div class="nextup">
        <div class="nextup-head">${nextHead}<a class="link-more" href="#schedule">Full schedule →</a></div>
        ${countdown}
        <div class="nextup-list">${nextList || '<p style="color:var(--muted)">No more matches on the schedule.</p>'}</div>
      </div>
    </section>

    <section class="stats">
      <div class="panel stat rise" style="--i:1">${ring(played / M.matches.length, 'var(--accent)')}<div><b>${played}<small> / ${M.matches.length}</small></b><p>Matches played</p></div></div>
      <div class="panel stat rise" style="--i:2">${ring(decided / M.events.length, 'var(--good)')}<div><b>${decided}<small> / ${M.events.length}</small></b><p>Events decided</p></div></div>
      <div class="panel stat rise" style="--i:3">${ring(1, 'var(--t3)')}<div><b>${playersN}</b><p>Players in individual events</p></div></div>
      <div class="panel stat rise" style="--i:4">${ring(span.first && span.last ? 1 - daysLeft / Math.max(1, (dateOfKey(keyOf(span.last)) - dateOfKey(keyOf(span.first))) / 864e5) : 0, 'var(--t4)')}<div><b>${daysLeft}</b><p>${daysLeft === 1 ? 'Day' : 'Days'} to finals night</p></div></div>
    </section>

    <section>
      <div class="section-head"><h2>The teams</h2><a class="link-more" href="#teams">All teams →</a></div>
      <div class="lineup">${[...M.teams].map((x, i) => `<a class="lineup-item rise" style="--i:${i};${tc(x)}" href="#team-${x.id}">
        <span class="lineup-logo"><img src="${logoSrc(x, true)}" alt="${esc(x.name)} logo" loading="lazy" decoding="async"><span class="ini">${initials(x.name)}</span></span>
        <span class="lineup-name"><b>${esc(x.name)}</b><small>${esc(x.code)} · ${x.points} pts</small></span></a>`).join('')}</div>
    </section>

    <div class="home-grid">
      <section>
        <div class="section-head"><h2>Team race</h2><a class="link-more" href="#standings">Standings →</a></div>
        <div class="panel board">
          ${ranked.map((x, i) => `<a class="board-row ${anyPts && x.rank === 1 ? 'lead' : ''}" href="#team-${x.id}" style="${tc(x)}">
            <span class="board-pos">${anyPts ? x.rank : i + 1}</span>${crest(x)}
            <span class="board-name"><b>${esc(x.name)}</b><span class="board-bar"><i style="width:${(x.points / maxPts) * 100}%"></i></span></span>
            <span class="board-pts"><b>${x.points}</b><small>pts</small></span></a>`).join('')}
        </div>
        ${anyPts ? '' : '<p style="color:var(--muted);font-size:.8rem;margin-top:10px">No points yet. Teams score 10 / 7 / 5 / 3 for 1st to 4th in each event.</p>'}
      </section>
      <section>
        <div class="section-head"><h2>Events</h2><a class="link-more" href="#brackets">All brackets →</a></div>
        <div class="events-grid">${M.events.map((e, i) => eventCard(e, i)).join('')}</div>
      </section>
    </div>`;
  }

  function eventCard(e, i) {
    const t = now();
    const champTeam = e.champion && M.teamOf(e.name, e.champion);
    const st = e.champion ? 'Decided' : e.played > 0 || (e.startAt && t >= e.startAt) ? 'In progress' : 'Starts ' + (e.startAt ? shortDate(e.startAt) : e.starts);
    const pct = e.matches.length ? (e.played / e.matches.length) * 100 : 0;
    return `<a class="panel ev-card rise" style="--i:${i}" href="#event-${e.slug}">
      <div class="ev-top"><span class="sport-ico">${icon(e.sport)}</span>
        <span class="ev-name"><b>${esc(e.name)}</b><small>${e.format === 'pools' ? '2 pools of 3 + knockout' : plural(e.entries, 'entry', 'entries')}${e.byes ? ` · ${plural(e.byes, 'bye')}` : ''}</small></span></div>
      ${e.champion ? `<div class="ev-champ" style="${tc(champTeam)}">${icon('trophy')}<span>${esc(e.champion)}</span>${champTeam ? `<i class="tdot" style="margin-left:auto"></i>` : ''}</div>` : `<div class="progress"><i style="width:${pct}%"></i></div>`}
      <div class="ev-foot"><span>${esc(st)}</span><span class="num">${e.played}/${e.matches.length} played</span></div>
    </a>`;
  }

  // ----- Schedule
  function viewSchedule(r) {
    const teams = M.teams;
    const days = groupDays(filteredMatches());
    const allDays = groupDays(M.matches);
    const todayKey = keyOf(now());
    const strip = allDays.map(([k, ms]) => {
      const d = dateOfKey(k);
      const finals = ms.some((m) => m.round === 'Final');
      return `<a href="#day-${k}" class="${k === todayKey ? 'today' : ''} ${finals ? 'finals' : ''} ${k < todayKey ? 'past' : ''}" title="${plural(ms.length, 'match', 'matches')}${finals ? ' · finals' : ''}"><small>${DAY[d.getDay()].slice(0, 3)}</small><b>${d.getDate()}</b><i></i></a>`;
    }).join('');
    return `
      <section>
        <div class="section-head"><div><span class="eyebrow">${plural(M.matches.length, 'match', 'matches')} · ${plural(allDays.length, 'day')}</span><h2 style="margin-top:6px">Schedule</h2></div>
          <p>Evenings 18:00–20:00 for Chess, Carrom and Table Tennis. Badminton runs all day on Sun 1 Nov. Cricket and Football start at 19:00.</p></div>
        <div style="display:flex;flex-direction:column;gap:14px">
          <nav class="datestrip" aria-label="Jump to a day">${strip}</nav>
          <div class="toolbar">
            <div class="seg" role="group" aria-label="Sport">${['All', ...SPORTS].map((s) => `<button type="button" data-sport="${s}" aria-pressed="${state.sport === s}">${s === 'All' ? '' : icon(s)}${s}</button>`).join('')}</div>
          </div>
          <div class="toolbar">
            <label class="field" for="f-q">${icon('users')}<input id="f-q" type="search" placeholder="Filter by player or match ID" value="${esc(state.q)}"></label>
            <label class="field" for="f-team"><select id="f-team"><option value="all">All teams</option>${teams.map((t) => `<option value="${t.id}" ${String(state.team) === String(t.id) ? 'selected' : ''}>${esc(t.name)}</option>`).join('')}</select></label>
            <label class="toggle" for="f-done"><input id="f-done" type="checkbox" ${state.hideDone ? 'checked' : ''}> Hide finished matches</label>
          </div>
        </div>
      </section>
      <section class="days" id="days">${renderDays(days, todayKey)}</section>`;
  }
  function groupDays(ms) {
    const map = new Map();
    ms.forEach((m) => { const k = m.startAt ? keyOf(m.startAt) : m.dayKey; if (!map.has(k)) map.set(k, []); map.get(k).push(m); });
    return [...map.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1));
  }
  function filteredMatches() {
    const q = state.q.trim().toLowerCase();
    return M.matches.filter((m) => {
      if (state.sport !== 'All' && sportOfMatch(m) !== state.sport) return false;
      if (state.hideDone && (m.winner || m.isDraw)) return false;
      if (state.team !== 'all') {
        const hit = m.slots.some((s) => { const t = M.teamOf(m.event, s.name); return t && String(t.id) === String(state.team); });
        if (!hit) return false;
      }
      if (q) {
        const hay = [m.id, m.event, m.round, ...m.slots.map((s) => s.name || '')].join(' ').toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }
  function renderDays(days, todayKey) {
    if (!days.length) return '<div class="empty">No matches match these filters.</div>';
    return days.map(([k, ms]) => {
      const d = dateOfKey(k);
      const finals = ms.some((m) => m.round === 'Final');
      const venues = [...new Set(ms.map((m) => m.venue).filter(Boolean))];
      return `<div class="day ${k < todayKey ? 'is-past' : ''}" id="d-${k}">
        <div class="day-head"><h3>${DAY[d.getDay()]} <span class="daynum">${d.getDate()}</span> ${MON[d.getMonth()]}</h3>
          ${k === todayKey ? '<span class="pill gold">Today</span>' : ''}${finals ? '<span class="pill gold">Finals</span>' : ''}
          <span>${plural(ms.length, 'match', 'matches')} · ${venues.join(', ')}</span></div>
        ${ms.map((m) => matchRow(m)).join('')}
      </div>`;
    }).join('');
  }

  // ----- Brackets
  function viewBrackets(r) {
    const ev = M.events.find((e) => e.slug === r.ev) || M.events[0];
    const bySport = SPORTS.map((s) => [s, M.events.filter((e) => e.sport === s)]).filter(([, l]) => l.length);
    const tabs = `<div class="ev-tabs">${bySport.map(([s, list]) => `<div class="ev-tabs-row"><span class="eyebrow">${s}</span>
      ${list.map((e) => `<a class="ev-tab ${e === ev ? 'active' : ''}" href="#event-${e.slug}">${icon(e.sport)}${esc(list.length > 1 || e.sport === 'Carrom' ? e.short.replace(/^Carroms /, '') : e.name)}</a>`).join('')}</div>`).join('')}</div>`;

    const places = [ev.champion, ev.runnerUp, ev.third];
    const podium = `<div class="podium" aria-label="Podium">${[1, 0, 2].map((i) => {
      const n = places[i]; const t = n && M.teamOf(ev.name, n);
      return `<div class="pod p${i + 1}" style="${tc(t)}">${t ? crest(t, 34) : ''}<span class="pod-name ${n ? '' : 'tbd'}">${n ? esc(n) : 'To be decided'}</span><div class="pod-step">${i + 1}</div></div>`;
    }).join('')}</div>`;

    const fdates = ev.startAt ? `${shortDate(ev.startAt)} – ${shortDate(ev.endAt)}` : '';
    const venues = [...new Set(ev.matches.map((m) => m.venue).filter(Boolean))].join(' · ');
    const head = `<section class="panel ev-hero">
      <div><span class="eyebrow">${esc(ev.sport)} · ${ev.format === 'pools' ? 'Pools + knockout' : 'Knockout + 3rd place match'}</span>
        <h2 style="margin-top:8px">${esc(ev.name)}</h2>
        <div class="chips">
          <span class="chip">${icon('users')}${ev.format === 'pools' ? '6 teams' : plural(ev.entries, 'entry', 'entries')}</span>
          ${ev.byes ? `<span class="chip">${plural(ev.byes, 'bye')}</span>` : ''}
          <span class="chip">${icon('calendar')}${esc(fdates)}</span>
          <span class="chip">${icon('pin')}${esc(venues)}</span>
          <span class="chip">${ev.played}/${ev.matches.length} played</span>
        </div></div>
      ${podium}
    </section>`;

    let body = '';
    if (ev.format === 'pools') body = poolsView(ev);
    body += `<section>
      <div class="section-head"><h2>${ev.format === 'pools' ? 'Knockout' : 'Bracket'}</h2>
        <div class="legend"><span><i class="tdot" style="background:var(--good)"></i>Winner advances</span><span><span class="pill live" style="padding:1px 7px">Live</span></span><span>Hover a name to trace their path</span></div></div>
      <div class="bracket-scroll">${bracketSVG(ev)}</div>
    </section>`;
    if (ev.format === 'knockout') {
      body += `<section><div class="section-head"><h2>Matches</h2><p>In schedule order. Byes start in the next round.</p></div>
        <div class="list">${ev.matches.map((m) => matchRow(m, { showDate: true })).join('')}</div></section>`;
    }
    return `<section>${tabs}</section>${head}${body}`;
  }

  function poolsView(ev) {
    const pools = Object.entries(ev.pools || {});
    return `<section>
      <div class="section-head"><h2>Pools</h2><p>Round robin inside each pool: ${ev.poolPoints.win} points for a win, ${ev.poolPoints.draw} for a draw or no result. Top two go through to the semi-finals.</p></div>
      ${ev.tieNotice ? '<p class="pill pending" style="margin-bottom:12px">A pool is tied on points. The organiser will set the tiebreak.</p>' : ''}
      <div class="pools">${pools.map(([p, rows]) => {
        const ms = ev.matches.filter((m) => m.pool === p || m.round === 'Pool ' + p);
        const complete = ms.length && ms.every((m) => m.winner || m.isDraw);
        return `<div class="panel pool"><h3>Pool ${p}</h3>
          <div class="tbl-wrap"><table class="tbl"><thead><tr><th>#</th><th>Team</th><th class="c">P</th><th class="c">W</th><th class="c">D</th><th class="c">L</th><th class="c">Pts</th></tr></thead>
          <tbody>${rows.map((x, i) => { const t = team(x.teamId); return `<tr style="${tc(t)}"><td class="num">${x.played ? x.rank : '–'}</td>
            <td><a class="team-cell" href="#team-${x.teamId}">${crest(t)}<span>${esc(x.name)}</span>${complete && x.rank <= 2 ? ' <span class="q">Q</span>' : ''}</a></td>
            <td class="c">${x.played}</td><td class="c">${x.won}</td><td class="c">${x.drawn}</td><td class="c">${x.lost}</td><td class="c pts">${x.points}</td></tr>`; }).join('')}</tbody></table></div>
          <div class="list">${ms.map((m) => matchRow(m, { showDate: true })).join('')}</div></div>`;
      }).join('')}</div></section>`;
  }

  // Bracket geometry: leaves (entries) are spaced evenly; each match sits at the mean of its two inputs.
  function bracketSVG(ev) {
    const byId = M.matchById;
    const final = ev.matches.find((m) => m.round === 'Final');
    if (!final) return '<div class="empty">No final found for this event.</div>';
    const third = ev.matches.find((m) => m.round === '3rd Place');
    const U = 50, W = 236, GAP = 58, HEAD = 24, SLOT = 32, CH = HEAD + SLOT * 2 + 2, TOP = 58;
    let leaf = 0; const nodes = [];
    (function walk(m, depth) {
      const kids = m.slots.map((s) => (s.src.kind === 'winner' && byId[s.src.match] && byId[s.src.match].event === m.event)
        ? walk(byId[s.src.match], depth + 1) : { leaf: true, pos: leaf++ });
      const node = { m, depth, kids, pos: (kids[0].pos + kids[1].pos) / 2 };
      nodes.push(node);
      return node;
    })(final, 0);
    const maxDepth = Math.max(...nodes.map((n) => n.depth));
    const colX = (n) => (maxDepth - n.depth) * (W + GAP);
    const cy = (n) => TOP + n.pos * U + U / 2;
    const finalNode = nodes.find((n) => n.m === final);
    const cols = maxDepth + 2;
    const width = cols * W + (cols - 1) * GAP;
    const thirdY = cy(finalNode) + CH + 50;
    let height = Math.max(TOP + leaf * U, thirdY + CH / 2 + 10, cy(finalNode) + 110);

    const roundName = (r) => ({ 'Quarter-final': 'Quarter-finals', 'Semi-final': 'Semi-finals' }[r] || r);
    const colHeads = [];
    for (let c = 0; c <= maxDepth; c++) {
      const ns = nodes.filter((n) => maxDepth - n.depth === c);
      const rounds = [...new Set(ns.map((n) => n.m.round))];
      const ds = ns.map((n) => n.m.startAt).filter(Boolean).sort((a, b) => a - b);
      const dl = ds.length ? (keyOf(ds[0]) === keyOf(ds[ds.length - 1]) ? shortDate(ds[0]) : `${shortDate(ds[0])} – ${shortDate(ds[ds.length - 1])}`) : '';
      colHeads.push(`<div class="col-head" style="left:${c * (W + GAP)}px;width:${W}px"><b>${esc(rounds.map(roundName).join(' / '))}</b><small>${plural(ns.length, 'match', 'matches')} · ${esc(dl)}</small></div>`);
    }
    colHeads.push(`<div class="col-head" style="left:${(maxDepth + 1) * (W + GAP)}px;width:${W}px"><b>Champion</b><small>${final.startAt ? shortDate(final.startAt) : ''}</small></div>`);

    const wires = [];
    const card = (m, x, yc, byeIdx) => {
      const st = status(m);
      const time = m.start && /^\d/.test(m.start) ? m.start : (m.start || '');
      const slots = m.slots.map((s, i) => {
        const t = M.teamOf(m.event, s.name);
        const cls = !s.name ? 'tbd' : m.winner ? (m.winner === s.name ? 'win' : 'lose') : '';
        return `<div class="slot ${cls}" ${s.name ? `data-entry="${esc(s.name)}"` : ''} style="${tc(t)}" title="${esc(slotLabel(s))}${t ? ' · ' + esc(t.name) : ''}">
          <i class="tdot"></i><span class="nm">${esc(slotLabel(s))}</span>${byeIdx && byeIdx.includes(i) ? '<span class="bye">BYE</span>' : ''}${t ? `<span class="tm">${initials(t.name)}</span>` : ''}</div>`;
      }).join('');
      return `<div class="mcard ${st === 'live' ? 'is-live' : ''} ${m.round === 'Final' ? 'is-final' : ''}" style="left:${x}px;top:${yc - CH / 2}px;width:${W}px">
        <div class="mc-head"><span>${esc(m.id)}</span><span>${st === 'live' ? '<span class="live-dot">● LIVE</span>' : esc(`${m.startAt ? shortDate(m.startAt) : ''}${time ? ' · ' + time : ''}`)}</span></div>${slots}</div>`;
    };

    const cards = nodes.map((n) => {
      const x = colX(n), y = cy(n);
      const byeIdx = n.depth < maxDepth ? n.kids.map((k, i) => (k.leaf && n.m.slots[i].src.kind === 'seed' ? i : -1)).filter((i) => i >= 0) : [];
      n.kids.forEach((k, i) => {
        if (k.leaf) return;
        const x1 = colX(k) + W, y1 = cy(k), x2 = x, y2 = y - CH / 2 + HEAD + SLOT * i + SLOT / 2 + 1, xm = x1 + GAP / 2;
        const wt = k.m.winner && M.teamOf(k.m.event, k.m.winner);
        wires.push(`<path class="${k.m.winner ? 'won' : ''}" ${wt ? `style="--wc:var(--t${wt.id})"` : ''} d="M${x1} ${y1} H${xm} V${y2} H${x2}"/>`);
      });
      return card(n.m, x, y, ev.format === 'knockout' ? byeIdx : null);
    });

    // champion + 3rd place
    const cx = (maxDepth + 1) * (W + GAP), fy = cy(finalNode);
    const champ = ev.champion, ct = champ && M.teamOf(ev.name, champ);
    wires.push(`<path class="${champ ? 'won' : ''}" ${ct ? `style="--wc:var(--t${ct.id})"` : ''} d="M${colX(finalNode) + W} ${fy} H${cx}"/>`);
    cards.push(`<div class="champ-card" style="left:${cx}px;top:${fy - 55}px;width:${W}px;height:110px;${tc(ct)}">${ct ? crest(ct, 40) : icon('trophy')}<b>${champ ? esc(champ) : 'To be decided'}</b><small>${ct ? (ct.name === champ ? 'Champion' : esc(ct.name)) : 'Winner of ' + esc(final.id)}</small></div>`);
    if (third) {
      const tx = colX(finalNode);
      cards.push(`<span class="tag-3p" style="left:${tx}px;top:${thirdY - CH / 2 - 26}px">3rd place match</span>`);
      cards.push(card(third, tx, thirdY));
      if (ev.third) {
        const t3 = M.teamOf(ev.name, ev.third);
        cards.push(`<div class="champ-card" style="left:${cx}px;top:${thirdY - 34}px;width:${W}px;height:68px;border-color:var(--bronze);background:var(--surface-2);${tc(t3)}"><small>3rd place</small><b style="font-size:1.05rem">${esc(ev.third)}</b></div>`);
        wires.push(`<path class="won" style="--wc:var(--bronze)" d="M${tx + W} ${thirdY} H${cx}"/>`);
      }
    }
    return `<div class="bracket" style="width:${width}px;height:${height}px">
      <svg class="wires" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" aria-hidden="true">${wires.join('')}</svg>
      ${colHeads.join('')}${cards.join('')}</div>`;
  }

  // ----- Standings
  function viewStandings() {
    const ranked = [...M.teams].sort((a, b) => (a.rank || 9) - (b.rank || 9) || b.points - a.points || a.id - b.id);
    const anyPts = ranked.some((t) => t.points > 0);
    const P = M.pts;
    const ordinal = ['1st', '2nd', '3rd', '4th'];
    const resultsRows = M.events.map((e) => {
      const names = [e.champion, e.runnerUp, e.third, e.fourth];
      return `<tr><td><a href="#event-${e.slug}" style="display:inline-flex;gap:8px;align-items:center">${icon(e.sport)}${esc(e.name)}</a></td>${names.map((n, i) => {
        let t = n ? M.teamOf(e.name, n) : null;
        if (!t && e.placeTeams && e.placeTeams[i]) t = M.teamOf(e.name, e.placeTeams[i]);
        return t ? `<td><span class="who" style="${tc(t)}"><i class="tdot"></i><span>${esc(n || t.name)}</span></span></td>` : '<td class="tbd">–</td>';
      }).join('')}</tr>`;
    }).join('');
    return `
      <section>
        <div class="section-head"><div><span class="eyebrow">Points by final position in every event</span><h2 style="margin-top:6px">Team standings</h2></div>
        <p>Singles points go to the player's team, doubles to the pair's team. Ties on points go to the team with more 1st places, then 2nd, then 3rd.</p></div>
        <div class="points-key">${[1, 2, 3, 4].map((i) => `<div class="pk p${i}"><span>${ordinal[i - 1]} place</span><b>${P[i]}</b></div>`).join('')}</div>
      </section>
      <section class="lb">
        ${ranked.map((t, i) => `<a class="lb-row rise ${anyPts && t.rank === 1 ? 'r1' : ''}" style="--i:${i};${tc(t)}" href="#team-${t.id}">
          <span class="lb-pos">${anyPts ? t.rank : '–'}</span>${crest(t)}
          <span class="lb-name"><b>${esc(t.name)}</b><small>${esc(t.code)}</small></span>
          <span class="medals">
            <span class="medal m1" title="1st places"><i></i>${t.firsts}</span><span class="medal m2" title="2nd places"><i></i>${t.seconds}</span>
            <span class="medal m3" title="3rd places"><i></i>${t.thirds}</span><span class="medal m4" title="4th places"><i></i>${t.fourths}</span></span>
          <span class="lb-pts"><b>${t.points}</b><small>points</small></span></a>`).join('')}
      </section>
      ${anyPts ? `<section><div class="section-head"><h2>Points breakdown</h2><p>How each team's total splits across finishing positions.</p></div>${pointsChart(ranked)}</section>` : ''}
      <section>
        <div class="section-head"><h2>Event results</h2><p>${M.events.filter((e) => e.champion).length} of ${M.events.length} events decided.</p></div>
        <div class="panel tbl-wrap"><table class="tbl results-tbl"><thead><tr><th>Event</th><th>1st · ${P[1]}</th><th>2nd · ${P[2]}</th><th>3rd · ${P[3]}</th><th>4th · ${P[4]}</th></tr></thead><tbody>${resultsRows}</tbody></table></div>
      </section>`;
  }

  function pointsChart(ranked) {
    const P = M.pts;
    const W = 760, rowH = 38, L = 190, R = 50, H = ranked.length * rowH + 30;
    const step = [1, 2, 5, 10, 20, 25, 50].find((st) => Math.max(1, ...ranked.map((t) => t.points)) / st <= 5) || 100;
    const max = Math.ceil(Math.max(1, ...ranked.map((t) => t.points)) / step) * step;
    const scale = (v) => (v / max) * (W - L - R);
    const fills = ['var(--accent)', 'var(--silver)', 'var(--bronze)', 'var(--muted)'];
    const ticks = Array.from({ length: max / step + 1 }, (_, i) => i * step);
    const rows = ranked.map((t, i) => {
      const y = 10 + i * rowH; let x = L;
      const segs = [t.firsts * P[1], t.seconds * P[2], t.thirds * P[3], t.fourths * P[4]].map((v, j) => {
        if (!v) return ''; const w = scale(v); const s = `<rect x="${x}" y="${y + 6}" width="${Math.max(0, w - 1.5)}" height="${rowH - 14}" rx="3" fill="${fills[j]}"><title>${['1st', '2nd', '3rd', '4th'][j]}: ${v} pts</title></rect>`; x += w; return s;
      }).join('');
      return `<text x="${L - 12}" y="${y + rowH / 2 + 4}" text-anchor="end">${esc(t.name)}</text>${segs}<text class="val" x="${x + 8}" y="${y + rowH / 2 + 5}">${t.points}</text>`;
    }).join('');
    const grid = ticks.map((v) => `<line class="grid" x1="${L + scale(v)}" x2="${L + scale(v)}" y1="6" y2="${H - 20}"/><text x="${L + scale(v)}" y="${H - 4}" text-anchor="middle">${v}</text>`).join('');
    return `<div class="panel chart tbl-wrap"><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Points by team, split by finishing position" style="min-width:560px">${grid}${rows}</svg>
      <div class="legend" style="margin-top:8px">${['1st', '2nd', '3rd', '4th'].map((l, j) => `<span><i class="tdot" style="background:${fills[j]};border-radius:3px"></i>${l} place</span>`).join('')}</div></div>`;
  }

  // ----- Teams
  function teamEntries(id) {
    return M.events.map((e) => [e, Object.entries(e.entryTeam).filter(([, tid]) => tid === id).map(([n]) => n)])
      .filter(([e, l]) => l.length && e.format === 'knockout');
  }
  function teamMatches(id) {
    return M.matches.filter((m) => m.slots.some((s) => { const t = M.teamOf(m.event, s.name); return t && t.id === id; }));
  }
  function viewTeams() {
    return `<section><div class="section-head"><div><span class="eyebrow">${plural(M.teams.length, 'team')}</span><h2 style="margin-top:6px">Teams</h2></div><p>Pick a team to see its players, entries and upcoming matches.</p></div>
      <div class="teams-grid">${[...M.teams].map((t, i) => {
        const players = M.players.filter((p) => p.teamId === t.id).length;
        const ents = teamEntries(t.id).reduce((a, [, l]) => a + l.length, 0);
        const next = teamMatches(t.id).find((m) => status(m) === 'upcoming' || status(m) === 'live');
        return `<a class="panel team-card rise" style="--i:${i};${tc(t)}" href="#team-${t.id}">
          <div class="row">${crest(t, 72)}<div><span class="eyebrow">${esc(t.code)}</span><h3>${esc(t.name)}</h3></div></div>
          <div class="kv"><div><b>${t.points}</b><small>Points</small></div><div><b>${players}</b><small>Players</small></div><div><b>${ents}</b><small>Entries</small></div></div>
          <p style="position:relative;z-index:1;color:var(--muted);font-size:.8rem">${next ? `Next: ${esc(next.event)} · ${shortDate(next.startAt)}${next.start && /^\d/.test(next.start) ? ' ' + next.start : ''}` : 'No matches left on the schedule'}</p>
        </a>`;
      }).join('')}</div></section>`;
  }

  function viewTeam(r) {
    const t = team(r.id);
    if (!t) return '<div class="empty">That team does not exist.</div>';
    const players = M.players.filter((p) => p.teamId === t.id);
    const ms = teamMatches(t.id);
    const upcoming = ms.filter((m) => !(m.winner || m.isDraw)).slice(0, 12);
    const results = ms.filter((m) => m.winner || m.isDraw).slice(-12).reverse();
    const ents = teamEntries(t.id);
    const ordinal = (n) => n + (['th', 'st', 'nd', 'rd'][(n % 100 - 20) % 10] || ['th', 'st', 'nd', 'rd'][n % 100] || 'th');
    return `
      <section class="team-hero rise" style="${tc(t)};--logo:url('${logoSrc(t, true)}')">${crest(t, 140, true)}
        <div class="grow"><span class="eyebrow">${esc(t.code)}</span><h2 style="margin-top:6px">${esc(t.name)}</h2>
          <div class="hero-meta" style="margin-top:14px"><span class="chip">${icon('users')}${plural(players.length, 'player')}</span><span class="chip">${plural(ents.reduce((a, [, l]) => a + l.length, 0), 'entry', 'entries')} in ${plural(ents.length, 'event')}</span>
          <span class="chip">${icon('Cricket')}Cricket &amp; Football squad</span></div></div>
        <div class="big"><b>${t.points}</b><small>Points · ${t.points ? ordinal(t.rank) + ' place' : 'no points yet'}</small>
          <div class="medals" style="justify-content:flex-end;margin-top:8px"><span class="medal m1"><i></i>${t.firsts}</span><span class="medal m2"><i></i>${t.seconds}</span><span class="medal m3"><i></i>${t.thirds}</span><span class="medal m4"><i></i>${t.fourths}</span></div></div>
      </section>
      <section><div class="section-head"><h2>Players</h2><p>Everyone entered in the individual and pairs events. Pick a name for their matches.</p></div>
        <div class="roster">${players.map((p) => `<a class="person" href="#player-${slug(p.name)}-${p.teamId}"><b>${esc(p.name)}</b><span class="tags">${p.events.map((e) => { const ev = M.eventByName[e]; return `<span class="mini" title="${esc(e)}">${ev ? icon(ev.sport) : ''}${esc(ev ? ev.short : e)}</span>`; }).join('')}</span></a>`).join('')}</div>
      </section>
      <div class="split">
        <section><div class="section-head"><h2>Coming up</h2><a class="link-more" href="#schedule" data-team-filter="${t.id}">In the schedule →</a></div>
          <div class="list">${upcoming.map((m) => matchRow(m, { showDate: true })).join('') || '<div class="empty">No matches left.</div>'}</div></section>
        <section><div class="section-head"><h2>Results</h2></div>
          <div class="list">${results.map((m) => matchRow(m, { showDate: true })).join('') || '<div class="empty">No results yet. They appear here as soon as the organiser records them.</div>'}</div></section>
      </div>`;
  }

  // ----- Player
  function findPlayer(key) { return M.players.find((p) => `${slug(p.name)}-${p.teamId}` === key); }
  function playerMatches(p) {
    return M.matches.filter((m) => p.events.includes(m.event) && m.slots.some((s) => s.name && s.name.split(/\s*&\s*/).includes(p.name) && (M.teamOf(m.event, s.name) || {}).id === p.teamId));
  }
  function viewPlayer(r) {
    const p = findPlayer(r.key);
    if (!p) return '<div class="empty">Player not found. Try the search.</div>';
    const t = team(p.teamId);
    const ms = playerMatches(p);
    const wins = ms.filter((m) => m.winner && m.winner.split(/\s*&\s*/).includes(p.name)).length;
    const losses = ms.filter((m) => m.loser && m.loser.split(/\s*&\s*/).includes(p.name)).length;
    const rows = ms.map((m) => {
      let note = '';
      if (!m.winner && !m.isDraw && m.round !== 'Final' && m.round !== '3rd Place') {
        const nxt = M.matches.find((x) => x.slots.some((s) => s.src.kind === 'winner' && s.src.match === m.id));
        if (nxt) note = `<p class="path-note">A win sends ${esc(p.name)} to ${esc(nxt.round)} ${esc(nxt.id)} · ${shortDate(nxt.startAt)}${nxt.start && /^\d/.test(nxt.start) ? ' at ' + nxt.start : ''}</p>`;
      }
      return matchRow(m, { showDate: true }) + note;
    }).join('');
    const medals = M.events.map((e) => {
      const places = [e.champion, e.runnerUp, e.third, e.fourth];
      const i = places.findIndex((n) => n && n.split(/\s*&\s*/).includes(p.name) && (M.teamOf(e.name, n) || {}).id === p.teamId);
      return i >= 0 ? `<span class="chip"><span class="medal m${i + 1}"><i></i></span>${['1st', '2nd', '3rd', '4th'][i]} · ${esc(e.name)}</span>` : '';
    }).join('');
    return `
      <section class="panel player-hero rise" style="${tc(t)}">
        <span class="avatar">${initials(p.name)}</span>
        <div style="flex:1;min-width:200px"><span class="eyebrow">Player</span><h2 style="margin-top:4px">${esc(p.name)}</h2>
          <div class="hero-meta" style="margin-top:10px"><a href="#team-${t.id}">${teamTag(t)}</a></div></div>
        <div class="kv" style="min-width:240px"><div><b>${p.events.length}</b><small>Events</small></div><div><b>${wins}</b><small>Wins</small></div><div><b>${losses}</b><small>Losses</small></div></div>
      </section>
      <section><div class="section-head"><h2>Events</h2></div><div class="hero-meta" style="margin-top:0">${p.events.map((e) => { const ev = M.eventByName[e]; return `<a class="ev-tab" href="#event-${ev.slug}">${icon(ev.sport)}${esc(e)}</a>`; }).join('')}${medals}</div></section>
      <section><div class="section-head"><h2>Matches</h2><p>${ms.length ? 'Later rounds appear once they qualify.' : ''}</p></div><div class="list">${rows || '<div class="empty">No matches scheduled.</div>'}</div></section>`;
  }

  // ---------- after render: bind controls
  let countdownTimer = null;
  function afterRender(r) {
    clearInterval(countdownTimer);
    const cd = $('#countdown');
    if (cd) {
      const at = +cd.dataset.at;
      const tick = () => {
        let s = Math.max(0, Math.floor((at - now()) / 1000));
        const d = Math.floor(s / 86400); s -= d * 86400; const h = Math.floor(s / 3600); s -= h * 3600; const m = Math.floor(s / 60); s -= m * 60;
        const set = (u, v) => { const el = cd.querySelector(`[data-u="${u}"]`); if (el) el.textContent = String(v).padStart(u === 'd' ? 1 : 2, '0'); };
        set('d', d); set('h', h); set('m', m); set('s', s);
        if (at - now() <= 0) { clearInterval(countdownTimer); render(true); }
      };
      tick(); countdownTimer = setInterval(tick, 1000);
    }
    if (r.name === 'schedule') {
      const redraw = () => { $('#days').innerHTML = renderDays(groupDays(filteredMatches()), keyOf(now())); };
      view.querySelectorAll('[data-sport]').forEach((b) => b.addEventListener('click', () => {
        state.sport = b.dataset.sport;
        view.querySelectorAll('[data-sport]').forEach((x) => x.setAttribute('aria-pressed', x === b));
        redraw();
      }));
      $('#f-q').addEventListener('input', (e) => { state.q = e.target.value; redraw(); });
      $('#f-team').addEventListener('change', (e) => { state.team = e.target.value; redraw(); });
      $('#f-done').addEventListener('change', (e) => { state.hideDone = e.target.checked; redraw(); });
    }
    view.querySelectorAll('[data-team-filter]').forEach((a) => a.addEventListener('click', () => { state.team = a.dataset.teamFilter; }));
    const br = view.querySelector('.bracket');
    if (br) {
      br.addEventListener('mouseover', (e) => {
        const s = e.target.closest('.slot[data-entry]');
        br.querySelectorAll('.slot.on').forEach((x) => x.classList.remove('on'));
        if (!s) { br.classList.remove('hl'); return; }
        br.classList.add('hl');
        br.querySelectorAll('.slot[data-entry]').forEach((x) => { if (x.dataset.entry === s.dataset.entry) x.classList.add('on'); });
      });
      br.addEventListener('mouseleave', () => { br.classList.remove('hl'); br.querySelectorAll('.slot.on').forEach((x) => x.classList.remove('on')); });
      // Open the bracket scrolled to the round being played now
      const live = br.querySelector('.mcard.is-live');
      if (live) live.scrollIntoView({ block: 'nearest', inline: 'center' });
    }
  }

  // ---------- search palette
  const pal = $('#palette'), palIn = $('#palette-input'), palList = $('#palette-results');
  let palItems = [], palSel = 0;
  function openPalette() {
    if (!M) return;
    pal.hidden = false; palIn.value = ''; searchPalette(''); setTimeout(() => palIn.focus(), 10);
  }
  function closePalette() { pal.hidden = true; }
  function searchPalette(q) {
    q = q.trim().toLowerCase();
    const items = [];
    M.players.forEach((p) => { const t = team(p.teamId); items.push({ kind: 'Player', title: p.name, sub: `${t ? t.name : ''} · ${p.events.join(', ')}`, href: `#player-${slug(p.name)}-${p.teamId}`, lead: `<span class="pal-avatar" style="${tc(t)}">${initials(p.name)}${t ? `<img class="pal-badge" src="${logoSrc(t)}" alt="" loading="lazy">` : ''}</span>`, hay: p.name }); });
    M.teams.forEach((t) => items.push({ kind: 'Team', title: t.name, sub: `${t.code} · ${t.points} pts`, href: `#team-${t.id}`, lead: crest(t), hay: `${t.name} ${t.code}` }));
    M.events.forEach((e) => items.push({ kind: 'Event', title: e.name, sub: `${e.sport} · ${e.played}/${e.matches.length} played`, href: `#event-${e.slug}`, lead: `<span class="sport-ico">${icon(e.sport)}</span>`, hay: `${e.name} ${e.sport} ${e.sheet}` }));
    M.matches.forEach((m) => items.push({ kind: 'Match', title: `${m.id} · ${m.slots.map(slotLabel).join(' vs ')}`, sub: `${m.event} · ${m.round} · ${m.startAt ? shortDate(m.startAt) : ''} ${m.start || ''}`, href: `#event-${(eventOf(m) || {}).slug}`, lead: `<span class="sport-ico">${icon(sportOfMatch(m))}</span>`, hay: m.id, matchOnly: true }));
    let res;
    if (!q) res = items.filter((x) => x.kind === 'Team' || x.kind === 'Event');
    else {
      res = items.filter((x) => (x.matchOnly ? x.hay.toLowerCase().includes(q) : x.hay.toLowerCase().includes(q)))
        .sort((a, b) => (b.hay.toLowerCase().startsWith(q) - a.hay.toLowerCase().startsWith(q)) || ['Player', 'Team', 'Event', 'Match'].indexOf(a.kind) - ['Player', 'Team', 'Event', 'Match'].indexOf(b.kind))
        .slice(0, 40);
    }
    palItems = res; palSel = 0;
    palList.innerHTML = res.length ? res.map((x, i) => `<li class="${i === 0 ? 'sel' : ''}" role="option"><a href="${x.href}">${x.lead}<span class="pal-text"><b>${esc(x.title)}</b><small>${esc(x.sub)}</small></span><em>${x.kind}</em></a></li>`).join('')
      : `<li class="none">No player, team, event or match matches “${esc(q)}”.</li>`;
  }
  function movePal(d) {
    const lis = palList.querySelectorAll('li[role="option"]'); if (!lis.length) return;
    lis[palSel].classList.remove('sel'); palSel = (palSel + d + lis.length) % lis.length; lis[palSel].classList.add('sel'); lis[palSel].scrollIntoView({ block: 'nearest' });
  }
  palIn.addEventListener('input', () => searchPalette(palIn.value));
  palIn.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') { e.preventDefault(); movePal(1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); movePal(-1); }
    else if (e.key === 'Enter') { const it = palItems[palSel]; if (it) { location.hash = it.href; closePalette(); } }
    else if (e.key === 'Escape') closePalette();
  });
  pal.addEventListener('click', (e) => { if (e.target.closest('[data-close]') || e.target.closest('a')) closePalette(); });
  $('#open-search').addEventListener('click', openPalette);
  document.addEventListener('keydown', (e) => {
    const typing = /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName);
    if ((e.key === '/' && !typing) || (e.key.toLowerCase() === 'k' && (e.metaKey || e.ctrlKey))) { e.preventDefault(); openPalette(); }
    else if (e.key === 'Escape' && !pal.hidden) closePalette();
  });

  // ---------- theme
  function applyTheme(t) { if (t) document.documentElement.setAttribute('data-theme', t); }
  try { applyTheme(localStorage.getItem('vc-theme')); } catch (e) { /* storage unavailable */ }
  $('#theme-toggle').addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark');
    const next = cur === 'dark' ? 'light' : 'dark';
    applyTheme(next);
    try { localStorage.setItem('vc-theme', next); } catch (e) { /* storage unavailable */ }
  });

  $('#file-input').addEventListener('change', onFilePick);
  document.addEventListener('error', (e) => {
    const img = e.target;
    if (img && img.tagName === 'IMG' && img.parentNode && (img.parentNode.classList.contains('has-logo') || img.parentNode.classList.contains('lineup-logo'))) {
      img.parentNode.classList.remove('has-logo'); img.remove();
    }
    else if (img && img.classList && img.classList.contains('pal-badge')) img.remove();
  }, true);
  window.addEventListener('hashchange', () => render());

  // ---------- boot
  let refreshTimer = null;
  loadRemote(true);
  refreshTimer = setInterval(() => { loadRemote(false).then(() => { if (M && lastSig !== 'local') setStatus(); }); }, CONFIG.refreshMs);
  // Keep live / upcoming states fresh even when the workbook hasn't changed
  setInterval(() => { if (M && !document.hidden && pal.hidden && route().name === 'home') render(true); }, 60 * 1000);
})();
