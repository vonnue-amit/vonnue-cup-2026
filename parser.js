/*
 * Vonnue Cup workbook parser.
 *
 * Reads Vonnue_Cup_Knockout_Brackets.xlsx (and optionally Vonnue_Cup_Participants.xlsx)
 * and builds the model the site renders. Bracket structure comes from the Schedule
 * sheet's formulas: every player cell is  =IF(<cell>="","Winner XX-R1-01",<cell>)
 * so the default text names the feeder match (or the seeded player for a bye),
 * and the calculated value shows who is actually there now. Workbooks saved without
 * calculated values are recalculated first (formula.js).
 */
(function (root) {
  const Formula = root.VonnueFormula || (typeof require === 'function' ? require('./formula.js') : null);
  const TEAM_RE = /^Team\s*(\d+)$/;
  const MATCH_ID_RE = /^[A-Z]+-[A-Z0-9]+-\d+$/;
  const PLACEHOLDER_RE = /^(Winner|Loser|W|L)\s+([A-Z]+-[A-Z0-9]+-\d+)$/;
  const POOL_SEED_RE = /^Pool [AB] (winner|runner-up)$/i;
  const MONTHS = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };

  function cellVal(ws, addr) {
    const c = ws && ws[addr];
    if (!c || c.v === undefined || c.v === null) return null;
    return typeof c.v === 'string' ? c.v.trim() : c.v;
  }

  function sheetRows(XLSX, ws) {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    const rows = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        row.push(cell ? { v: cell.v === undefined ? null : (typeof cell.v === 'string' ? cell.v.trim() : cell.v), f: cell.f || null, addr } : { v: null, f: null, addr });
      }
      rows.push(row);
    }
    return rows;
  }

  function isPlaceholder(s) {
    if (s == null || s === '' || s === '—') return true;
    return PLACEHOLDER_RE.test(s) || POOL_SEED_RE.test(s) || /^Tie\b/.test(s) || /^L final$/.test(s);
  }

  // Default text in =IF(x="","<default>",x): the feeder / seed for that slot.
  function slotSource(cell) {
    let def = null;
    if (cell.f) {
      const m = cell.f.match(/=\s*""\s*,\s*"([^"]*)"/);
      if (m) def = m[1].trim();
    }
    if (def == null && typeof cell.v === 'string') def = cell.v;
    if (def == null) return { kind: 'empty' };
    const p = def.match(PLACEHOLDER_RE);
    if (p) return { kind: p[1][0] === 'W' ? 'winner' : 'loser', match: p[2] };
    if (POOL_SEED_RE.test(def)) return { kind: 'pool', label: def };
    return { kind: 'seed', name: def };
  }

  function parseDate(str, year) {
    const m = String(str || '').match(/(\d{1,2})\s+([A-Z][a-z]{2})/);
    if (!m) return null;
    return { y: year, m: MONTHS[m[2]], d: +m[1] };
  }
  function toDate(dp, hhmm) {
    if (!dp) return null;
    const t = String(hhmm || '').match(/^(\d{1,2}):(\d{2})$/);
    return new Date(dp.y, dp.m, dp.d, t ? +t[1] : 0, t ? +t[2] : 0);
  }

  function sportOf(eventName) {
    if (/^Chess/i.test(eventName)) return 'Chess';
    if (/^Carrom/i.test(eventName)) return 'Carrom';
    if (/^TT /.test(eventName)) return 'Table Tennis';
    if (/^Badminton/.test(eventName)) return 'Badminton';
    if (/^Cricket/.test(eventName)) return 'Cricket';
    if (/^Football/.test(eventName)) return 'Football';
    return 'Other';
  }

  function shortEvent(name) {
    return name.replace(/^Badminton /, '').replace(/^TT /, '').replace(" Doubles", ' Doubles');
  }

  // Read a workbook and fill in any formula results the file was saved without.
  function readWorkbook(XLSX, buf) {
    const wb = buf && buf.Sheets ? buf : XLSX.read(buf, { type: buf instanceof ArrayBuffer ? 'array' : undefined, cellFormula: true, sheetStubs: true, cellNF: false, cellStyles: false });
    wb.recalculated = Formula ? Formula.recalcMissing(wb) : 0;
    return wb;
  }

  function parse(XLSX, bracketsBuf, participantsBuf, opts) {
    opts = opts || {};
    const year = opts.year || 2026;
    const wb = readWorkbook(XLSX, bracketsBuf);
    const S = (n) => wb.Sheets[n];

    // ---- Teams (Standings "Points by team" block)
    const teams = [];
    const teamById = {};
    const teamByName = {};
    const standings = S('Standings');
    const stRows = sheetRows(XLSX, standings);
    let pts = { 1: 10, 2: 7, 3: 5, 4: 3 };
    stRows.forEach((row) => {
      const a = row[0].v;
      if (a === '1st' && typeof row[1].v === 'number') {
        pts = { 1: row[1].v, 2: row[3].v, 3: row[5].v, 4: row[7].v };
      }
      if (typeof a === 'string' && TEAM_RE.test(a) && typeof row[1].v === 'string' && typeof row[6].v === 'number') {
        const id = +a.match(TEAM_RE)[1];
        const t = {
          id, code: a, name: row[1].v,
          firsts: row[2].v || 0, seconds: row[3].v || 0, thirds: row[4].v || 0, fourths: row[5].v || 0,
          points: row[6].v || 0, rank: row[7].v || null,
        };
        teams.push(t); teamById[id] = t; teamByName[t.name.toLowerCase()] = t;
      }
    });
    teams.sort((a, b) => a.id - b.id);

    // ---- Events (Overview)
    const events = [];
    const eventByName = {};
    sheetRows(XLSX, S('Overview')).forEach((row) => {
      const [name, sheet, entries, matches, byes, starts, final, champ, runner, third] = row.map((c) => c.v);
      if (typeof name === 'string' && typeof sheet === 'string' && wb.Sheets[sheet] && typeof entries === 'number') {
        const ev = {
          name, sheet, entries, matchCount: matches, byes, starts, finalDate: final,
          sport: sportOf(name), short: shortEvent(name),
          slug: sheet.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
          champion: champ && champ !== '—' ? champ : null,
          runnerUp: runner && runner !== '—' ? runner : null,
          third: third && third !== '—' ? third : null,
          fourth: null, matches: [], format: /Cricket|Football/.test(name) ? 'pools' : 'knockout',
        };
        events.push(ev); eventByName[name] = ev;
      }
    });

    // Event results by team (Standings "Event results": 1st..4th team names)
    stRows.forEach((row) => {
      const ev = eventByName[row[0].v];
      if (ev && row.length > 4) {
        ev.placeTeams = [1, 2, 3, 4].map((i) => (row[i].v ? String(row[i].v) : null));
      }
    });

    // ---- Team of each entry, per event, from the bracket sheet: "Team N" cell followed by a name
    events.forEach((ev) => {
      ev.entryTeam = {};
      const ws = S(ev.sheet);
      sheetRows(XLSX, ws).forEach((row) => {
        for (let i = 0; i < row.length - 1; i++) {
          const a = row[i].v, b = row[i + 1].v;
          if (typeof a === 'string' && TEAM_RE.test(a) && typeof b === 'string' && b && !isPlaceholder(b) && !TEAM_RE.test(b)) {
            if (!ev.entryTeam[b]) ev.entryTeam[b] = +a.match(TEAM_RE)[1];
          }
        }
      });
    });
    events.filter((e) => e.format === 'pools').forEach((ev) => {
      ev.entryTeam = {};
      teams.forEach((t) => { ev.entryTeam[t.name] = t.id; });
    });

    // ---- Schedule
    const matches = [];
    const matchById = {};
    let currentDay = null;
    sheetRows(XLSX, S('Schedule')).forEach((row) => {
      const id = row[6].v;
      if (typeof id === 'string' && MATCH_ID_RE.test(id)) {
        const evName = row[4].v;
        const ev = eventByName[evName] || events.find((e) => e.sheet === evName) || null;
        const dateStr = row[0].v;
        const dp = parseDate(dateStr, year);
        const start = row[1].v, end = row[2].v;
        const m = {
          id, event: ev ? ev.name : evName, round: row[5].v, venue: row[3].v,
          dateLabel: dateStr, dayKey: dp ? `${dp.y}-${String(dp.m + 1).padStart(2, '0')}-${String(dp.d).padStart(2, '0')}` : dateStr,
          start: start || null, end: end || null,
          startAt: toDate(dp, start), endAt: /^\d{1,2}:\d{2}$/.test(String(end || '')) ? toDate(dp, end) : null,
          dayNote: currentDay,
          slots: [row[7], row[8]].map((c) => {
            const v = typeof c.v === 'string' ? c.v : (c.v == null ? null : String(c.v));
            const src = slotSource(c);
            const known = v && !isPlaceholder(v) && !(src.kind === 'pool' && v === src.label);
            return { name: known ? v : null, label: v, src };
          }),
          winner: null, loser: null, score: null, isDraw: false,
        };
        matches.push(m); matchById[id] = m;
        if (ev) ev.matches.push(m);
      } else if (typeof row[0].v === 'string' && row[1].v == null && row[6].v == null && /\d/.test(row[0].v) && /·|—/.test(row[0].v)) {
        currentDay = row[0].v;
      }
    });

    // ---- Pool sheets (Cricket, Football): tables, scores, winners, semi-final seeds
    events.filter((e) => e.format === 'pools').forEach((ev) => {
      const ws = S(ev.sheet);
      const rows = sheetRows(XLSX, ws);
      ev.pools = {};
      ev.poolPoints = { win: null, draw: null };
      let pool = null, hdr = null, matchHdr = null;
      rows.forEach((row) => {
        const a = row[0].v;
        if (a === 'Points: Win') ev.poolPoints = { win: row[1].v, draw: row[3].v };
        const pm = typeof a === 'string' && a.match(/^Pool ([A-Z])$/);
        if (pm) { pool = pm[1]; ev.pools[pool] = []; return; }
        if (a === 'Team' && row[1].v === 'Team Name') {
          hdr = {}; row.forEach((c, i) => { if (typeof c.v === 'string' && i < 10) hdr[c.v] = i; }); return;
        }
        if (a === 'Match ID') { matchHdr = {}; row.forEach((c, i) => { if (typeof c.v === 'string' && i < 10) matchHdr[c.v] = i; }); pool = null; return; }
        if (pool && hdr && typeof a === 'string' && TEAM_RE.test(a)) {
          const g = (k) => (hdr[k] != null ? row[hdr[k]].v : null);
          ev.pools[pool].push({
            teamId: +a.match(TEAM_RE)[1], name: g('Team Name'), played: g('Played') || 0, won: g('Won') || 0,
            drawn: g('Drawn') || 0, lost: g('Lost') || 0, points: g('Points') || 0, tiebreak: g('Tiebreak'), rank: g('Rank'),
          });
        }
        if (matchHdr && typeof a === 'string' && matchById[a]) {
          const m = matchById[a];
          const score = matchHdr['Score / Result'] != null ? row[matchHdr['Score / Result']].v : null;
          const win = matchHdr['Winner'] != null ? row[matchHdr['Winner']].v : null;
          m.pool = matchHdr['Pool'] != null ? row[matchHdr['Pool']].v : null;
          if (score != null && score !== '') m.score = String(score);
          if (win) {
            const w = String(win).trim();
            const names = m.slots.map((s) => s.name);
            if (/^draw|no result/i.test(w)) m.isDraw = true;
            else if (w === 'Team 1' || w === names[0]) { m.winner = names[0]; m.loser = names[1]; }
            else if (w === 'Team 2' || w === names[1]) { m.winner = names[1]; m.loser = names[0]; }
          }
        }
      });
      // Football goals for / against, from scores written like "2-1" (Team 1's goals first)
      if (ev.sport === 'Football') {
        Object.values(ev.pools).forEach((p) => p.forEach((x) => { x.gf = 0; x.ga = 0; }));
        const rowOf = {};
        Object.values(ev.pools).forEach((p) => p.forEach((x) => { rowOf[x.name] = x; }));
        ev.matches.forEach((m) => {
          const g = m.pool && m.score && m.score.match(/^\s*(\d+)\s*[-–—:]\s*(\d+)/);
          if (!g) return;
          const [a, b] = m.slots.map((s) => rowOf[s.name]);
          if (a) { a.gf += +g[1]; a.ga += +g[2]; }
          if (b) { b.gf += +g[2]; b.ga += +g[1]; }
        });
        Object.values(ev.pools).forEach((p) => p.forEach((x) => { x.gd = x.gf - x.ga; }));
      }
      Object.values(ev.pools).forEach((p) => p.sort((x, y) => (x.rank || 9) - (y.rank || 9) || y.points - x.points));
      // Pool-stage seed labels for the knockout (e.g. a tie still to be resolved)
      ev.tieNotice = matches.some((m) => m.event === ev.name && m.slots.some((s) => /^Tie\b/.test(s.label || '')));
    });

    // ---- Knockout results: a match is decided once its winner / loser appears in a later slot
    matches.forEach((m) => {
      m.slots.forEach((s) => {
        if (!s.name) return;
        const feeder = (s.src.kind === 'winner' || s.src.kind === 'loser') && matchById[s.src.match];
        if (!feeder) return;
        const other = (n) => feeder.slots.map((x) => x.name).find((x) => x && x !== n) || null;
        if (s.src.kind === 'winner') { feeder.winner = s.name; feeder.loser = feeder.loser || other(s.name); }
        else { feeder.loser = s.name; feeder.winner = feeder.winner || other(s.name); }
      });
    });
    events.forEach((ev) => {
      const final = ev.matches.find((m) => m.round === 'Final');
      const third = ev.matches.find((m) => m.round === '3rd Place');
      const pick = (m, w) => {
        if (!m || !w) return;
        const names = m.slots.map((s) => s.name);
        if (names.includes(w)) { m.winner = w; m.loser = names.find((n) => n && n !== w) || null; }
      };
      pick(final, ev.champion);
      pick(third, ev.third);
      if (third && third.winner) ev.fourth = third.loser;
      ev.played = ev.matches.filter((m) => m.winner || m.isDraw).length;
      ev.startAt = ev.matches.reduce((a, m) => (m.startAt && (!a || m.startAt < a) ? m.startAt : a), null);
      ev.endAt = ev.matches.reduce((a, m) => (m.startAt && (!a || m.startAt > a) ? m.startAt : a), null);
    });

    // Resolve the team of any slot name
    function teamOf(eventName, name) {
      if (!name) return null;
      const ev = eventByName[eventName];
      const id = ev && ev.entryTeam[name];
      if (id) return teamById[id] || null;
      const tm = String(name).match(TEAM_RE);
      if (tm) return teamById[+tm[1]] || null;
      return teamByName[String(name).toLowerCase()] || null;
    }

    // ---- Participants workbook (rosters)
    let entries = [];
    if (participantsBuf) {
      try {
        const pwb = readWorkbook(XLSX, participantsBuf);
        const ws = pwb.Sheets['All Entries'];
        if (ws) {
          XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }).slice(1).forEach((r) => {
            if (!r || !r[1] || !r[2]) return;
            const tm = String(r[1]).match(/Team\s*(\d+)/);
            entries.push({
              teamId: tm ? +tm[1] : null, sport: r[2], division: r[3], event: r[4],
              players: [r[5], r[6]].filter(Boolean).map((x) => String(x).trim()),
            });
          });
        }
      } catch (e) { entries = []; }
    }

    // ---- Players index (from the brackets so names match the schedule exactly)
    const players = {};
    function addPlayer(name, teamId, eventName, entryName) {
      const key = `${name.toLowerCase()}|${teamId || 0}`;
      if (!players[key]) players[key] = { key, name, teamId, events: new Set(), entryNames: new Set() };
      players[key].events.add(eventName);
      players[key].entryNames.add(entryName);
    }
    events.filter((e) => e.format === 'knockout').forEach((ev) => {
      Object.entries(ev.entryTeam).forEach(([entry, tid]) => {
        entry.split(/\s*&\s*/).forEach((p) => addPlayer(p, tid, ev.name, entry));
      });
    });
    const playerList = Object.values(players).map((p) => ({ ...p, events: [...p.events], entryNames: [...p.entryNames] }))
      .sort((a, b) => a.name.localeCompare(b.name));

    return {
      year, pts, teams, teamById, events, eventByName, matches, matchById, entries, players: playerList, teamOf,
      loadedAt: new Date(), recalculated: wb.recalculated,
    };
  }

  const api = { parse, readWorkbook, isPlaceholder, slotSource };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.VonnueParser = api;
})(typeof window !== 'undefined' ? window : globalThis);
