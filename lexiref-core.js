/*!
 * LexiRef Core — citation engine shared by the LexiRef Windows app and the LexiRef Word add-in.
 * Pure JavaScript, no dependencies. Works in browsers and in Node (for tests).
 *
 * Exposes a global `LexiRef` namespace:
 *   LexiRef.TYPES, LexiRef.STYLES, LexiRef.FIELDS
 *   LexiRef.newItem(type), LexiRef.normalizeItem(item)
 *   LexiRef.parseNames(text), LexiRef.namesToText(names)
 *   LexiRef.formatReference(item, style, index) -> { html, text }
 *   LexiRef.formatCitation(items, style, opts)  -> { html, text }
 *   LexiRef.sortForBibliography(items, style)
 *   LexiRef.bibliography(items, style, opts)   -> { entries:[{item, html, text, n}], html, text }
 *   LexiRef.bibtex.parse(text) / LexiRef.bibtex.serialize(items)
 *   LexiRef.ris.parse(text)    / LexiRef.ris.serialize(items)
 *   LexiRef.csl.toCSL(item) / LexiRef.csl.fromCSL(obj) / LexiRef.csl.parse(text)
 *   LexiRef.crossref.lookupDOI(doi), LexiRef.crossref.search(query), LexiRef.integrity.check(item)
 *   LexiRef.exportDocx(entries, opts) -> Blob, LexiRef.exportRtf(entries, opts) -> string
 *   LexiRef.findDuplicates(items), LexiRef.citeKey(item), LexiRef.libraryFile.serialize / parse
 */
(function (root) {
  'use strict';

  /* ------------------------------------------------------------------ */
  /*  Reference types and fields                                         */
  /* ------------------------------------------------------------------ */

  var TYPES = {
    journalArticle:  { label: 'Journal article',       icon: 'article',  container: 'Journal' },
    book:            { label: 'Book',                  icon: 'book',     container: 'Series' },
    bookSection:     { label: 'Book chapter',          icon: 'chapter',  container: 'Book title' },
    conferencePaper: { label: 'Conference paper',      icon: 'conf',     container: 'Proceedings / conference' },
    thesis:          { label: 'Thesis / dissertation', icon: 'thesis',   container: '' },
    report:          { label: 'Report',                icon: 'report',   container: 'Series' },
    webpage:         { label: 'Web page',              icon: 'web',      container: 'Website name' },
    preprint:        { label: 'Preprint',              icon: 'preprint', container: 'Repository (e.g. arXiv)' },
    dataset:         { label: 'Dataset',               icon: 'dataset',  container: '' },
    patent:          { label: 'Patent',                icon: 'patent',   container: '' },
    standard:        { label: 'Standard',              icon: 'standard', container: '' },
    software:        { label: 'Software',              icon: 'software', container: '' },
    other:           { label: 'Other',                 icon: 'other',    container: 'Source / container' }
  };

  // Field definitions (key -> label, input kind). Labels may be overridden per type.
  var FIELDS = {
    title:          { label: 'Title',               kind: 'text' },
    authors:        { label: 'Authors',             kind: 'names' },
    editors:        { label: 'Editors',             kind: 'names' },
    year:           { label: 'Year',                kind: 'number' },
    month:          { label: 'Month',               kind: 'month' },
    day:            { label: 'Day',                 kind: 'number' },
    containerTitle: { label: 'Journal',             kind: 'text' },
    volume:         { label: 'Volume',              kind: 'text' },
    issue:          { label: 'Issue',               kind: 'text' },
    pages:          { label: 'Pages',               kind: 'text' },
    publisher:      { label: 'Publisher',           kind: 'text' },
    place:          { label: 'Place',               kind: 'text' },
    edition:        { label: 'Edition',             kind: 'text' },
    series:         { label: 'Series',              kind: 'text' },
    institution:    { label: 'Institution',         kind: 'text' },
    genre:          { label: 'Type / genre',        kind: 'text' },
    number:         { label: 'Number',              kind: 'text' },
    version:        { label: 'Version',             kind: 'text' },
    isbn:           { label: 'ISBN',                kind: 'text' },
    issn:           { label: 'ISSN',                kind: 'text' },
    doi:            { label: 'DOI',                 kind: 'text' },
    url:            { label: 'URL',                 kind: 'url' },
    accessed:       { label: 'Accessed (date)',     kind: 'date' },
    language:       { label: 'Language',            kind: 'text' },
    abstract:       { label: 'Abstract',            kind: 'textarea' },
    tags:           { label: 'Tags (comma separated)', kind: 'tags' },
    attachment:     { label: 'Attachment / file link', kind: 'text' }
  };

  // Which fields the editor shows for each type, with label overrides.
  var TYPE_FIELDS = {
    journalArticle:  ['title','authors','year','month','containerTitle:Journal','volume','issue','pages','doi','url','issn','language','abstract','tags','attachment'],
    book:            ['title','authors','editors','year','edition','publisher','place','series','volume','isbn','doi','url','language','abstract','tags','attachment'],
    bookSection:     ['title','authors','editors','year','containerTitle:Book title','edition','pages','publisher','place','isbn','doi','url','language','abstract','tags','attachment'],
    conferencePaper: ['title','authors','editors','year','month','day','containerTitle:Proceedings / conference','place','publisher','volume','pages','doi','url','language','abstract','tags','attachment'],
    thesis:          ['title','authors','year','month','genre:Thesis type (e.g. PhD thesis)','institution:University','place','doi','url','language','abstract','tags','attachment'],
    report:          ['title','authors','year','month','day','institution:Institution / agency','number:Report number','series','publisher','place','doi','url','language','abstract','tags','attachment'],
    webpage:         ['title','authors','year','month','day','containerTitle:Website name','url','accessed','language','abstract','tags','attachment'],
    preprint:        ['title','authors','year','month','day','containerTitle:Repository (e.g. arXiv)','number:Identifier (e.g. arXiv:2010.11929)','doi','url','language','abstract','tags','attachment'],
    dataset:         ['title','authors','year','month','day','publisher:Repository / publisher','version','doi','url','accessed','language','abstract','tags','attachment'],
    patent:          ['title','authors:Inventors','year','month','day','number:Patent number','institution:Assignee','place:Country / patent office','url','language','abstract','tags','attachment'],
    standard:        ['title','authors:Organization','year','number:Standard number','publisher','place','doi','url','language','abstract','tags','attachment'],
    software:        ['title','authors','year','version','publisher','place','url','doi','language','abstract','tags','attachment'],
    other:           ['title','authors','editors','year','month','day','containerTitle:Source / container','volume','issue','pages','publisher','place','institution','number','doi','url','accessed','language','abstract','tags','attachment']
  };

  var STYLES = {
    apa:       { label: 'APA 7th edition',           short: 'APA',       numeric: false },
    ieee:      { label: 'IEEE',                      short: 'IEEE',      numeric: true },
    mla:       { label: 'MLA 9th edition',           short: 'MLA',       numeric: false },
    chicago:   { label: 'Chicago (author–date)',     short: 'Chicago',   numeric: false },
    harvard:   { label: 'Harvard (Cite Them Right)', short: 'Harvard',   numeric: false },
    vancouver: { label: 'Vancouver',                 short: 'Vancouver', numeric: true }
  };

  var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var MONTHS_ABBR = ['Jan.','Feb.','Mar.','Apr.','May','June','July','Aug.','Sept.','Oct.','Nov.','Dec.'];
  var MONTHS_IEEE = ['Jan.','Feb.','Mar.','Apr.','May','Jun.','Jul.','Aug.','Sep.','Oct.','Nov.','Dec.'];
  var MONTHS_3 = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  /* ------------------------------------------------------------------ */
  /*  Small utilities                                                    */
  /* ------------------------------------------------------------------ */

  function uid(prefix) {
    var s = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    return (prefix || 'r_') + s;
  }
  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function stripHtml(h) {
    return String(h || '').replace(/<[^>]+>/g, '')
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&nbsp;/g, ' ');
  }
  function trim(s) { return String(s == null ? '' : s).trim(); }
  function has(v) { return v != null && String(v).trim() !== ''; }
  function endsWithPunct(s) { return /[.!?…]["'’”)\]]*\s*$/.test(stripHtml(s)); }
  function period(s) { s = trim(s); if (!s) return ''; return endsWithPunct(s) ? s : s + '.'; }
  function enDash(p) { return String(p || '').replace(/\s*(--|-|—|–)\s*/g, '–'); }
  function hyphen(p) { return String(p || '').replace(/\s*(--|—|–)\s*/g, '-'); }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function join(parts, sep) {
    var out = [];
    for (var i = 0; i < parts.length; i++) if (has(parts[i])) out.push(trim(parts[i]));
    return out.join(sep);
  }
  function sentenceEnd(parts) { // join parts that each end with a period
    var out = [];
    for (var i = 0; i < parts.length; i++) if (has(parts[i])) out.push(period(parts[i]));
    return out.join(' ');
  }
  function doiUrl(doi) { doi = cleanDoi(doi); return doi ? 'https://doi.org/' + doi : ''; }
  function cleanDoi(doi) {
    doi = trim(doi);
    if (!doi) return '';
    doi = doi.replace(/^(https?:\/\/)?(dx\.)?doi\.org\//i, '').replace(/^doi:\s*/i, '');
    return doi;
  }
  function yearOf(item) { return has(item.year) ? String(item.year) : 'n.d.'; }
  function monthName(m, arr) { m = parseInt(m, 10); return (m >= 1 && m <= 12) ? (arr || MONTHS)[m - 1] : ''; }
  function parseDateStr(s) { // "2024-05-17" | "2024/05/17" | "17 May 2024" | "May 17, 2024" -> {y,m,d}
    s = trim(s); if (!s) return null;
    var m = s.match(/^(\d{4})[-\/.](\d{1,2})(?:[-\/.](\d{1,2}))?/);
    if (m) return { y: +m[1], m: +m[2], d: m[3] ? +m[3] : null };
    var d = new Date(s);
    if (!isNaN(d.getTime())) return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
    var y = s.match(/\d{4}/); return y ? { y: +y[0], m: null, d: null } : null;
  }
  function fmtAccessed(s, style) {
    var d = parseDateStr(s); if (!d) return '';
    var mName = d.m ? monthName(d.m) : '';
    switch (style) {
      case 'apa': case 'chicago': return (mName ? mName + (d.d ? ' ' + d.d : '') + ', ' : '') + d.y;
      case 'mla': return (d.d ? d.d + ' ' : '') + (d.m ? monthName(d.m, MONTHS_ABBR) + ' ' : '') + d.y;
      case 'ieee': return (d.m ? monthName(d.m, MONTHS_IEEE) + ' ' : '') + (d.d ? d.d + ', ' : '') + d.y;
      case 'harvard': return (d.d ? d.d + ' ' : '') + (mName ? mName + ' ' : '') + d.y;
      case 'vancouver': return d.y + (d.m ? ' ' + monthName(d.m, MONTHS_3) : '') + (d.d ? ' ' + d.d : '');
    }
    return s;
  }
  function todayISO() { var d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }

  /* ------------------------------------------------------------------ */
  /*  Items                                                              */
  /* ------------------------------------------------------------------ */

  function newItem(type) {
    var now = new Date().toISOString();
    return {
      id: uid('r_'), type: type && TYPES[type] ? type : 'journalArticle',
      title: '', authors: [], editors: [], year: '', month: '', day: '',
      containerTitle: '', volume: '', issue: '', pages: '', publisher: '', place: '', edition: '', series: '',
      institution: '', genre: '', number: '', version: '', isbn: '', issn: '', doi: '', url: '', accessed: '',
      language: '', abstract: '', tags: [], collections: [], notes: '', attachment: '',
      favorite: false, retracted: false, retractionNote: '', integrity: { status: 'unchecked', checked: '' },
      added: now, modified: now
    };
  }

  function normalizeItem(raw) {
    var it = newItem(raw && raw.type);
    if (!raw) return it;
    for (var k in raw) if (Object.prototype.hasOwnProperty.call(raw, k)) it[k] = raw[k];
    if (!it.id) it.id = uid('r_');
    if (!TYPES[it.type]) it.type = 'other';
    it.authors = normNames(it.authors);
    it.editors = normNames(it.editors);
    if (!Array.isArray(it.tags)) it.tags = has(it.tags) ? String(it.tags).split(/[,;]/).map(trim).filter(Boolean) : [];
    if (!Array.isArray(it.collections)) it.collections = [];
    if (!it.integrity || typeof it.integrity !== 'object') it.integrity = { status: 'unchecked', checked: '' };
    it.doi = cleanDoi(it.doi);
    if (has(it.year)) { var y = String(it.year).match(/\d{4}/); it.year = y ? +y[0] : ''; }
    if (!it.added) it.added = new Date().toISOString();
    if (!it.modified) it.modified = it.added;
    return it;
  }
  function normNames(list) {
    if (!list) return [];
    if (typeof list === 'string') return parseNames(list);
    var out = [];
    for (var i = 0; i < list.length; i++) {
      var n = list[i];
      if (!n) continue;
      if (typeof n === 'string') { out = out.concat(parseNames(n)); continue; }
      if (n.literal) { out.push({ literal: trim(n.literal) }); continue; }
      var fam = trim(n.family || n.last || ''), giv = trim(n.given || n.first || '');
      if (fam || giv) out.push({ family: fam, given: giv });
    }
    return out;
  }

  // Parse "Family, Given" / "Given Family" lines; "{Organization Name}" => literal.
  function parseNames(text) {
    var out = [];
    String(text || '').split(/\r?\n|;/).forEach(function (line) {
      // Also split " and " / " & " (BibTeX style) when the line has no newline structure
      line.split(/\s+and\s+|\s*&\s*/).forEach(function (p) {
        p = trim(p); if (!p) return;
        var lit = p.match(/^\{(.+)\}$/);
        if (lit) { out.push({ literal: trim(lit[1]) }); return; }
        if (p.indexOf(',') >= 0) {
          var parts = p.split(','), fam = trim(parts[0]), giv = trim(parts.slice(1).join(' '));
          // "Smith, Jr., John" -> family "Smith, Jr." given "John"
          var m = p.match(/^(.+?),\s*(Jr\.?|Sr\.?|II|III|IV),\s*(.+)$/i);
          if (m) { fam = trim(m[1]) + ', ' + trim(m[2]); giv = trim(m[3]); }
          out.push({ family: fam, given: giv });
        } else {
          var w = p.split(/\s+/);
          if (w.length === 1) out.push({ family: w[0], given: '' });
          else {
            // keep particles with family: "Ludwig van Beethoven" -> family "van Beethoven"
            var idx = w.length - 1;
            while (idx > 1 && /^(van|von|de|del|della|der|den|da|di|du|le|la|al|el|bin|ibn|ter|ten|zu)$/i.test(w[idx - 1])) idx--;
            out.push({ family: w.slice(idx).join(' '), given: w.slice(0, idx).join(' ') });
          }
        }
      });
    });
    return out;
  }
  function namesToText(names) {
    return (names || []).map(function (n) {
      if (n.literal) return '{' + n.literal + '}';
      return n.given ? n.family + ', ' + n.given : n.family;
    }).join('\n');
  }
  function nameDisplay(n) { if (!n) return ''; if (n.literal) return n.literal; return join([n.given, n.family], ' '); }

  /* ------------------------------------------------------------------ */
  /*  Name formatting                                                    */
  /* ------------------------------------------------------------------ */

  function initials(given, opts) {
    opts = opts || {};
    var dots = opts.dots !== false, sep = opts.sep != null ? opts.sep : ' ';
    given = trim(given); if (!given) return '';
    var parts = given.split(/\s+/), out = [];
    parts.forEach(function (p) {
      var hy = p.split('-'), seg = [];
      hy.forEach(function (h) {
        h = h.replace(/[.,]/g, '');
        if (!h) return;
        seg.push(h.charAt(0).toUpperCase() + (dots ? '.' : ''));
      });
      if (seg.length) out.push(seg.join(dots ? '-' : '-'));
    });
    return out.join(sep);
  }
  // "Vaswani, A."
  function nameInvInit(n, opts) {
    if (n.literal) return n.literal;
    var ini = initials(n.given, opts);
    return n.family + (ini ? ', ' + ini : '');
  }
  // "A. Vaswani"
  function nameInitFirst(n) {
    if (n.literal) return n.literal;
    var ini = initials(n.given);
    return (ini ? ini + ' ' : '') + n.family;
  }
  // "Vaswani, Ashish"
  function nameInvFull(n) { if (n.literal) return n.literal; return n.family + (n.given ? ', ' + n.given : ''); }
  // "Ashish Vaswani"
  function nameFull(n) { if (n.literal) return n.literal; return join([n.given, n.family], ' '); }
  // Vancouver: "Vaswani A"
  function nameVanc(n) { if (n.literal) return n.literal; var ini = initials(n.given, { dots: false, sep: '' }); return n.family + (ini ? ' ' + ini : ''); }
  function family(n) { return n.literal ? n.literal : n.family; }

  function authorsAPA(names, opts) {
    opts = opts || {};
    var f = function (n) { return nameInvInit(n); };
    if (!names.length) return '';
    if (names.length === 1) return f(names[0]);
    if (names.length === 2) return f(names[0]) + ', & ' + f(names[1]);
    if (names.length <= 20) return names.slice(0, -1).map(f).join(', ') + ', & ' + f(names[names.length - 1]);
    return names.slice(0, 19).map(f).join(', ') + ', . . . ' + f(names[names.length - 1]);
  }
  function editorsAPA(names) { // "A. B. Smith & C. Jones"
    var f = function (n) { return nameInitFirst(n); };
    if (!names.length) return '';
    if (names.length === 1) return f(names[0]);
    if (names.length === 2) return f(names[0]) + ' & ' + f(names[1]);
    return names.slice(0, -1).map(f).join(', ') + ', & ' + f(names[names.length - 1]);
  }
  function authorsIEEE(names) {
    var f = nameInitFirst;
    if (!names.length) return '';
    if (names.length === 1) return f(names[0]);
    if (names.length === 2) return f(names[0]) + ' and ' + f(names[1]);
    if (names.length <= 6) return names.slice(0, -1).map(f).join(', ') + ', and ' + f(names[names.length - 1]);
    return f(names[0]) + ' et al.';
  }
  function authorsMLA(names) {
    if (!names.length) return '';
    if (names.length === 1) return nameInvFull(names[0]);
    if (names.length === 2) return nameInvFull(names[0]) + ', and ' + nameFull(names[1]);
    return nameInvFull(names[0]) + ', et al.';
  }
  function authorsChicago(names) {
    if (!names.length) return '';
    if (names.length === 1) return nameInvFull(names[0]);
    if (names.length <= 10) {
      var arr = [nameInvFull(names[0])].concat(names.slice(1).map(nameFull));
      if (arr.length === 2) return arr[0] + ', and ' + arr[1];
      return arr.slice(0, -1).join(', ') + ', and ' + arr[arr.length - 1];
    }
    return [nameInvFull(names[0])].concat(names.slice(1, 7).map(nameFull)).join(', ') + ', et al.';
  }
  function authorsHarvard(names) {
    var f = function (n) { return nameInvInit(n); };
    if (!names.length) return '';
    if (names.length === 1) return f(names[0]);
    if (names.length === 2) return f(names[0]) + ' and ' + f(names[1]);
    if (names.length === 3) return f(names[0]) + ', ' + f(names[1]) + ' and ' + f(names[2]);
    return f(names[0]) + ' et al.';
  }
  function authorsVancouver(names) {
    if (!names.length) return '';
    if (names.length <= 6) return names.map(nameVanc).join(', ');
    return names.slice(0, 6).map(nameVanc).join(', ') + ', et al';
  }
  function editorsNatural(names, mla) { // "Daniel Kahneman and Amos Tversky"
    var f = nameFull;
    if (!names.length) return '';
    if (names.length === 1) return f(names[0]);
    if (names.length === 2) return f(names[0]) + ' and ' + f(names[1]);
    if (mla && names.length > 2) return f(names[0]) + ' et al.';
    return names.slice(0, -1).map(f).join(', ') + ', and ' + f(names[names.length - 1]);
  }

  /* ------------------------------------------------------------------ */
  /*  Reference list entry formatting                                    */
  /* ------------------------------------------------------------------ */

  function it_(s) { return has(s) ? '<i>' + esc(s) + '</i>' : ''; }
  function q(s, open, close) { return has(s) ? open + esc(s) + close : ''; }
  function titlePeriod(s) { return has(s) ? period(esc(s)) : ''; }

  function thesisGenre(item, style) {
    var g = trim(item.genre).toLowerCase();
    var phd = !g || /ph\.?\s*d|doctor|dissertation|dr\.|habil/.test(g);
    var master = /master|m\.?\s*sc|m\.?\s*a\b|msc|m\.eng|m\.s\./.test(g);
    var bachelor = /bachelor|b\.?\s*sc|b\.?\s*a\b|bsc/.test(g);
    if (g && !phd && !master && !bachelor) return trim(item.genre);
    switch (style) {
      case 'apa':       return bachelor ? "Bachelor's thesis" : master ? "Master's thesis" : 'Doctoral dissertation';
      case 'ieee':      return bachelor ? 'B.S. thesis' : master ? 'M.S. thesis' : 'Ph.D. dissertation';
      case 'mla':       return bachelor ? 'Bachelor\u2019s thesis' : master ? 'Master\u2019s thesis' : 'PhD dissertation';
      case 'chicago':   return bachelor ? 'Bachelor\u2019s thesis' : master ? 'Master\u2019s thesis' : 'PhD diss.';
      case 'harvard':   return bachelor ? 'Bachelor\u2019s thesis' : master ? 'Master\u2019s thesis' : 'PhD thesis';
      case 'vancouver': return bachelor ? 'bachelor\u2019s thesis' : master ? 'master\u2019s thesis' : 'dissertation';
    }
    return item.genre || 'Thesis';
  }
  function ordinal(ed) {
    ed = trim(ed); if (!ed) return '';
    var n = parseInt(ed, 10);
    if (isNaN(n)) return ed; // "Revised"
    var s = ['th', 'st', 'nd', 'rd'], v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
  }
  function apaDate(item) {
    var y = yearOf(item);
    if (has(item.month) && (item.type === 'webpage' || item.type === 'conferencePaper' || item.type === 'preprint' || item.type === 'report' || item.type === 'dataset' || item.type === 'patent')) {
      return y + ', ' + monthName(item.month) + (has(item.day) ? ' ' + item.day : '');
    }
    return y;
  }
  function locator(item, style) {
    var doi = cleanDoi(item.doi);
    if (doi) {
      if (style === 'ieee') return 'doi: ' + esc(doi) + '.';
      if (style === 'harvard') return 'doi:' + esc(doi) + '.';
      if (style === 'vancouver') return 'doi:' + esc(doi);
      return esc(doiUrl(doi));
    }
    if (has(item.url)) {
      if (style === 'ieee') return '[Online]. Available: ' + esc(item.url);
      if (style === 'harvard') return 'Available at: ' + esc(item.url) + (has(item.accessed) ? ' (Accessed: ' + fmtAccessed(item.accessed, 'harvard') + ').' : '.');
      if (style === 'vancouver') return 'Available from: ' + esc(item.url);
      return esc(item.url);
    }
    return '';
  }

  var FORMATTERS = {};

  /* ---- APA 7 ---- */
  FORMATTERS.apa = function (item) {
    var A = authorsAPA(item.authors), E = item.editors || [];
    var date = '(' + apaDate(item) + ').';
    var head = A ? period(A) + ' ' + date : '';
    var loc = locator(item, 'apa');
    var out = [];
    function titleFirst(titleHtml) { // no author: title moves to author position
      return titleHtml + ' ' + date;
    }
    switch (item.type) {
      case 'journalArticle': case 'other': {
        var vol = has(item.volume) ? ', ' + it_(item.volume) : '';
        var iss = has(item.issue) ? '(' + esc(item.issue) + ')' : '';
        var pg = has(item.pages) ? ', ' + esc(enDash(item.pages)) : '';
        var src = has(item.containerTitle) ? it_(item.containerTitle) + vol + iss + pg + '.' : '';
        out = A ? [head, titlePeriod(item.title), src, loc] : [titleFirst(titlePeriod(item.title)), src, loc];
        break;
      }
      case 'book': {
        var ed = has(item.edition) ? ' (' + ordinal(item.edition) + ' ed.)' : '';
        var t = it_(item.title) + ed + '.';
        if (!A && E.length) { A = authorsAPA(E) + (E.length > 1 ? ' (Eds.)' : ' (Ed.)'); head = period(A) + ' ' + date; }
        out = A ? [head, t, period(esc(item.publisher)), loc] : [titleFirst(t), period(esc(item.publisher)), loc];
        break;
      }
      case 'bookSection': {
        var eds = E.length ? 'In ' + editorsAPA(E) + (E.length > 1 ? ' (Eds.), ' : ' (Ed.), ') : 'In ';
        var det = [];
        if (has(item.edition)) det.push(ordinal(item.edition) + ' ed.');
        if (has(item.pages)) det.push('pp. ' + enDash(item.pages));
        var src2 = eds + it_(item.containerTitle) + (det.length ? ' (' + esc(det.join(', ')) + ')' : '') + '.';
        out = [head, titlePeriod(item.title), src2, period(esc(item.publisher)), loc];
        break;
      }
      case 'conferencePaper': {
        if (has(item.containerTitle)) {
          var eds2 = E.length ? 'In ' + editorsAPA(E) + (E.length > 1 ? ' (Eds.), ' : ' (Ed.), ') : 'In ';
          var det2 = [];
          if (has(item.volume)) det2.push('Vol. ' + item.volume);
          if (has(item.pages)) det2.push('pp. ' + enDash(item.pages));
          out = [head, titlePeriod(item.title), eds2 + it_(item.containerTitle) + (det2.length ? ' (' + esc(det2.join(', ')) + ')' : '') + '.', period(esc(item.publisher)), loc];
        } else {
          out = [head, it_(item.title) + ' [Paper presentation].', period(esc(join([item.publisher, item.place], ', '))), loc];
        }
        break;
      }
      case 'thesis': {
        var br = ' [' + thesisGenre(item, 'apa') + (has(item.institution) ? ', ' + esc(item.institution) : '') + '].';
        out = [head, it_(item.title) + br, has(item.publisher) ? period(esc(item.publisher)) : '', loc];
        break;
      }
      case 'report': {
        var num = has(item.number) ? ' (' + esc(/report|no\.|rep\./i.test(item.number) ? item.number : 'Report No. ' + item.number) + ')' : '';
        var pub = esc(item.institution || item.publisher);
        if (A && A === pub) pub = '';
        out = A ? [head, it_(item.title) + num + '.', period(pub), loc] : [titleFirst(it_(item.title) + num + '.'), period(pub), loc];
        break;
      }
      case 'webpage': {
        var site = esc(item.containerTitle);
        if (A && A === site) site = '';
        var retrieved = has(item.accessed) ? 'Retrieved ' + fmtAccessed(item.accessed, 'apa') + ', from ' + esc(item.url) : esc(item.url);
        out = A ? [head, it_(item.title) + '.', period(site), retrieved] : [titleFirst(it_(item.title) + '.'), period(site), retrieved];
        break;
      }
      case 'preprint': {
        out = [head, it_(item.title) + '.', period(esc(item.containerTitle)), loc];
        break;
      }
      case 'dataset': {
        var ver = has(item.version) ? ' (Version ' + esc(item.version) + ')' : '';
        out = [head, it_(item.title) + ver + ' [Data set].', period(esc(item.publisher)), loc];
        break;
      }
      case 'patent': {
        var pn = has(item.number) ? ' (' + esc(join([item.place, 'Patent No. ' + item.number], ' ')) + ')' : '';
        out = [head, it_(item.title) + pn + '.', period(esc(item.institution || item.publisher)), loc];
        break;
      }
      case 'standard': {
        var sn = has(item.number) ? ' (' + esc(item.number) + ')' : '';
        var pub2 = esc(item.publisher || item.institution);
        if (A && A === pub2) pub2 = '';
        out = A ? [head, it_(item.title) + sn + '.', period(pub2), loc] : [titleFirst(it_(item.title) + sn + '.'), period(pub2), loc];
        break;
      }
      case 'software': {
        var v2 = has(item.version) ? ' (Version ' + esc(item.version) + ')' : '';
        out = [head, it_(item.title) + v2 + ' [Computer software].', period(esc(item.publisher)), loc];
        break;
      }
    }
    return join(out, ' ');
  };

  /* ---- IEEE ---- */
  FORMATTERS.ieee = function (item) {
    var A = authorsIEEE(item.authors), E = item.editors || [];
    var y = yearOf(item), mo = has(item.month) ? monthName(item.month, MONTHS_IEEE) : '';
    var loc = locator(item, 'ieee');
    var locSep = cleanDoi(item.doi) ? ',' : '.'; // ", doi: ..." but ". [Online]. Available: ..."
    var t = '\u201C' + esc(item.title) + ',\u201D';
    var out = [];
    switch (item.type) {
      case 'journalArticle': case 'other': {
        var parts = [it_(item.containerTitle)];
        if (has(item.volume)) parts.push('vol. ' + esc(item.volume));
        if (has(item.issue)) parts.push('no. ' + esc(item.issue));
        if (has(item.pages)) parts.push((item.pages.indexOf('-') >= 0 || item.pages.indexOf('–') >= 0 ? 'pp. ' : 'p. ') + esc(enDash(item.pages)));
        parts.push(join([mo, y], ' '));
        out = [A ? A + ',' : '', t, join(parts, ', ') + (loc ? locSep : '.'), loc];
        break;
      }
      case 'book': {
        if (!A && E.length) A = authorsIEEE(E) + (E.length > 1 ? ', Eds.' : ', Ed.');
        var ed = has(item.edition) ? ', ' + ordinal(item.edition) + ' ed' : '';
        out = [A ? A + ',' : '', it_(item.title) + ed + '.', join([item.place ? esc(item.place) + ':' : '', esc(item.publisher)], ' ') + ', ' + y + '.', loc];
        break;
      }
      case 'bookSection': {
        var eds = E.length ? ', ' + authorsIEEE(E) + (E.length > 1 ? ', Eds.' : ', Ed.') : '';
        var pub = join([item.place ? esc(item.place) + ':' : '', esc(item.publisher)], ' ');
        out = [A ? A + ',' : '', t, 'in ' + it_(item.containerTitle) + eds + '.', join([pub, y, has(item.pages) ? 'pp. ' + esc(enDash(item.pages)) : ''], ', ') + '.', loc];
        break;
      }
      case 'conferencePaper': {
        var tail = join([esc(item.place), join([mo, y], ' '), has(item.pages) ? 'pp. ' + esc(enDash(item.pages)) : ''], ', ');
        out = [A ? A + ',' : '', t, 'in ' + it_(item.containerTitle) + (tail ? ', ' + tail : '') + (loc ? locSep : '.'), loc];
        break;
      }
      case 'thesis': {
        out = [A ? A + ',' : '', t, join([thesisGenre(item, 'ieee'), esc(item.institution), esc(item.place), y], ', ') + '.', loc];
        break;
      }
      case 'report': {
        var num = has(item.number) ? esc(/rep|no\./i.test(item.number) ? item.number : 'Rep. ' + item.number) : '';
        out = [A ? A + ',' : '', t, join([esc(item.institution || item.publisher), esc(item.place), num, join([mo, y], ' ')], ', ') + '.', loc];
        break;
      }
      case 'webpage': {
        var acc = has(item.accessed) ? ' (accessed ' + fmtAccessed(item.accessed, 'ieee') + ').' : '';
        out = [A ? A + ',' : '', '\u201C' + esc(item.title) + '.\u201D', period(esc(item.containerTitle)), has(item.url) ? esc(item.url) + acc : ''];
        break;
      }
      case 'preprint': {
        out = [A ? A + ',' : '', t, join([y, has(item.number) ? it_(item.number) : it_(item.containerTitle)], ', ') + '.', loc];
        break;
      }
      case 'dataset': {
        out = [A ? A + ',' : '', t, join([esc(item.publisher), y], ', ') + '. [Dataset].', loc];
        break;
      }
      case 'patent': {
        out = [A ? A + ',' : '', t, join([esc(item.place), 'Patent ' + esc(item.number)], ' ') + ', ' + join([mo, has(item.day) ? item.day + ',' : '', y], ' ').replace(',,', ',') + '.', loc];
        break;
      }
      case 'standard': {
        out = [it_(item.title) + ',', join([esc(item.number), esc(item.publisher || item.institution || A), esc(item.place), y], ', ') + '.', loc];
        break;
      }
      case 'software': {
        out = [A ? A + ',' : '', esc(item.title) + (has(item.version) ? ' (Version ' + esc(item.version) + ')' : '') + '. (' + y + ').', '[Software].', loc];
        break;
      }
    }
    return join(out, ' ');
  };

  /* ---- MLA 9 ---- */
  FORMATTERS.mla = function (item) {
    var A = authorsMLA(item.authors), E = item.editors || [];
    var y = has(item.year) ? String(item.year) : '';
    var date = join([has(item.day) ? String(item.day) : '', has(item.month) ? monthName(item.month, MONTHS_ABBR) : '', y], ' ');
    var tq = '\u201C' + period(esc(item.title)) + '\u201D';
    var ti = it_(item.title) + '.';
    var doi = cleanDoi(item.doi);
    var loc = doi ? esc(doiUrl(doi)) + '.' : has(item.url) ? esc(item.url.replace(/^https?:\/\//, '')) + '.' : '';
    var acc = has(item.accessed) ? 'Accessed ' + fmtAccessed(item.accessed, 'mla') + '.' : '';
    var out = [];
    switch (item.type) {
      case 'journalArticle': case 'other': {
        var c = [it_(item.containerTitle)];
        if (has(item.volume)) c.push('vol. ' + esc(item.volume));
        if (has(item.issue)) c.push('no. ' + esc(item.issue));
        if (date) c.push(date);
        if (has(item.pages)) c.push('pp. ' + esc(enDash(item.pages)));
        out = [A ? period(A) : '', tq, join(c, ', ') + '.', loc];
        break;
      }
      case 'book': {
        if (!A && E.length) A = authorsMLA(E) + (E.length > 1 ? ', editors' : ', editor');
        var c2 = [];
        if (has(item.edition)) c2.push(ordinal(item.edition) + ' ed.');
        if (has(item.publisher)) c2.push(esc(item.publisher));
        if (y) c2.push(y);
        out = [A ? period(A) : '', ti, join(c2, ', ') + '.', loc];
        break;
      }
      case 'bookSection': case 'conferencePaper': {
        var c3 = [it_(item.containerTitle)];
        if (E.length) c3.push('edited by ' + esc(editorsNatural(E, true)));
        if (has(item.edition)) c3.push(ordinal(item.edition) + ' ed.');
        if (has(item.publisher)) c3.push(esc(item.publisher));
        if (y) c3.push(y);
        if (has(item.pages)) c3.push('pp. ' + esc(enDash(item.pages)));
        out = [A ? period(A) : '', tq, join(c3, ', ') + '.', loc];
        break;
      }
      case 'thesis': {
        out = [A ? period(A) : '', ti, join([y, esc(item.institution)], '. ') + ', ' + thesisGenre(item, 'mla') + '.', loc];
        break;
      }
      case 'report': {
        out = [A ? period(A) : '', ti, join([esc(item.institution || item.publisher), y], ', ') + '.', loc];
        break;
      }
      case 'webpage': {
        var site = esc(item.containerTitle);
        if (A && stripHtml(A) === item.containerTitle) site = '';
        out = [A ? period(A) : '', tq, join([site ? it_(item.containerTitle) : '', date], ', ') + (has(item.url) ? ', ' + esc(item.url.replace(/^https?:\/\//, '')) : '') + '.', acc];
        break;
      }
      case 'preprint': {
        out = [A ? period(A) : '', tq, join([it_(item.containerTitle), esc(item.number), date], ', ') + '.', loc];
        break;
      }
      case 'dataset': {
        out = [A ? period(A) : '', ti, join([esc(item.publisher), date], ', ') + '.', loc];
        break;
      }
      case 'patent': {
        out = [A ? period(A) : '', ti, join([esc(join([item.place, 'Patent ' + item.number], ' ')), date], ', ') + '.', loc];
        break;
      }
      case 'standard': {
        out = [A ? period(A) : '', ti, join([esc(item.number), esc(item.publisher || item.institution), y], ', ') + '.', loc];
        break;
      }
      case 'software': {
        out = [A ? period(A) : '', ti, join([has(item.version) ? 'Version ' + esc(item.version) : '', esc(item.publisher), y], ', ') + '.', loc];
        break;
      }
    }
    return join(out, ' ');
  };

  /* ---- Chicago author-date ---- */
  FORMATTERS.chicago = function (item) {
    var A = authorsChicago(item.authors), E = item.editors || [];
    var y = yearOf(item);
    var tq = '\u201C' + period(esc(item.title)) + '\u201D';
    var ti = it_(item.title) + '.';
    var doi = cleanDoi(item.doi);
    var loc = doi ? esc(doiUrl(doi)) + '.' : has(item.url) ? esc(item.url) + '.' : '';
    var pubPlace = join([esc(item.place), esc(item.publisher)], ': ');
    var out = [];
    switch (item.type) {
      case 'journalArticle': case 'other': {
        var src = it_(item.containerTitle) + (has(item.volume) ? ' ' + esc(item.volume) : '') + (has(item.issue) ? ' (' + esc(item.issue) + ')' : '') + (has(item.pages) ? ': ' + esc(enDash(item.pages)) : '') + '.';
        out = [A ? period(A) : '', y + '.', tq, src, loc];
        break;
      }
      case 'book': {
        if (!A && E.length) A = authorsChicago(E) + (E.length > 1 ? ', eds' : ', ed');
        out = [A ? period(A) : '', y + '.', ti, has(item.edition) ? ordinal(item.edition) + ' ed.' : '', pubPlace ? pubPlace + '.' : '', loc];
        break;
      }
      case 'bookSection': case 'conferencePaper': {
        var src2 = 'In ' + it_(item.containerTitle) + (E.length ? ', edited by ' + esc(editorsNatural(E)) : '') + (has(item.pages) ? ', ' + esc(enDash(item.pages)) : '') + '.';
        out = [A ? period(A) : '', y + '.', tq, src2, pubPlace ? pubPlace + '.' : '', loc];
        break;
      }
      case 'thesis': {
        out = [A ? period(A) : '', y + '.', tq, join([thesisGenre(item, 'chicago'), esc(item.institution)], ', ') + '.', loc];
        break;
      }
      case 'report': {
        out = [A ? period(A) : '', y + '.', ti, has(item.number) ? esc(item.number) + '.' : '', join([esc(item.place), esc(item.institution || item.publisher)], ': ') + '.', loc];
        break;
      }
      case 'webpage': {
        var dt = has(item.month) ? monthName(item.month) + (has(item.day) ? ' ' + item.day : '') + ', ' + y : '';
        var site = esc(item.containerTitle);
        if (A && stripHtml(A) === item.containerTitle) site = '';
        out = [A ? period(A) : '', y + '.', tq, site ? site + '.' : '', dt ? dt + '.' : '', has(item.accessed) ? 'Accessed ' + fmtAccessed(item.accessed, 'chicago') + '.' : '', has(item.url) ? esc(item.url) + '.' : ''];
        break;
      }
      case 'preprint': {
        out = [A ? period(A) : '', y + '.', tq, join([esc(item.containerTitle), esc(item.number)], ', ') + '.', loc];
        break;
      }
      case 'dataset': {
        out = [A ? period(A) : '', y + '.', ti, has(item.version) ? 'Version ' + esc(item.version) + '.' : '', esc(item.publisher) + '.', loc];
        break;
      }
      case 'patent': {
        out = [A ? period(A) : '', y + '.', ti, esc(join([item.place, 'Patent ' + item.number], ' ')) + '.', loc];
        break;
      }
      case 'standard': {
        out = [A ? period(A) : '', y + '.', ti, esc(item.number) + '.', pubPlace ? pubPlace + '.' : '', loc];
        break;
      }
      case 'software': {
        out = [A ? period(A) : '', y + '.', ti, has(item.version) ? 'Version ' + esc(item.version) + '.' : '', esc(item.publisher) + '.', loc];
        break;
      }
    }
    return join(out, ' ');
  };

  /* ---- Harvard (Cite Them Right) ---- */
  FORMATTERS.harvard = function (item) {
    var A = authorsHarvard(item.authors), E = item.editors || [];
    var y = '(' + yearOf(item) + ')';
    var tq = '\u2018' + esc(item.title) + '\u2019';
    var ti = it_(item.title);
    var loc = locator(item, 'harvard');
    var pubPlace = join([esc(item.place), esc(item.publisher)], ': ');
    var out = [];
    switch (item.type) {
      case 'journalArticle': case 'other': {
        var src = it_(item.containerTitle) + (has(item.volume) ? ', ' + esc(item.volume) : '') + (has(item.issue) ? '(' + esc(item.issue) + ')' : '') + (has(item.pages) ? ', pp. ' + esc(enDash(item.pages)) : '') + '.';
        out = [A, y, tq + ',', src, loc];
        break;
      }
      case 'book': {
        if (!A && E.length) A = authorsHarvard(E) + (E.length > 1 ? ' (eds.)' : ' (ed.)');
        out = [A, y, ti + '.', has(item.edition) ? ordinal(item.edition) + ' edn.' : '', pubPlace ? pubPlace + '.' : '', loc];
        break;
      }
      case 'bookSection': case 'conferencePaper': {
        var eds = E.length ? 'in ' + authorsHarvard(E) + (E.length > 1 ? ' (eds.) ' : ' (ed.) ') : 'in ';
        out = [A, y, tq + ',', eds + it_(item.containerTitle) + '.', join([pubPlace, has(item.pages) ? 'pp. ' + esc(enDash(item.pages)) : ''], ', ') + '.', loc];
        break;
      }
      case 'thesis': {
        out = [A, y, ti + '.', thesisGenre(item, 'harvard') + '.', period(esc(item.institution)), loc];
        break;
      }
      case 'report': {
        out = [A, y, ti + '.', join([esc(item.place), esc(item.institution || item.publisher)], ': ') + (has(item.number) ? ' (' + esc(item.number) + ')' : '') + '.', loc];
        break;
      }
      case 'webpage': {
        out = [A, y, ti + '.', has(item.url) ? 'Available at: ' + esc(item.url) + (has(item.accessed) ? ' (Accessed: ' + fmtAccessed(item.accessed, 'harvard') + ').' : '.') : ''];
        break;
      }
      case 'preprint': {
        out = [A, y, tq + '.', join([esc(item.containerTitle), esc(item.number)], ', ') + ' [Preprint].', loc];
        break;
      }
      case 'dataset': {
        out = [A, y, ti + ' [Dataset].', period(esc(item.publisher)), loc];
        break;
      }
      case 'patent': {
        out = [A, y, ti + '.', esc(join([item.place, 'Patent no. ' + item.number], ' ')) + '.', loc];
        break;
      }
      case 'standard': {
        out = [A, y, join([esc(item.number), ti], ': ') + '.', pubPlace ? pubPlace + '.' : '', loc];
        break;
      }
      case 'software': {
        out = [A, y, ti + (has(item.version) ? ' (Version ' + esc(item.version) + ')' : '') + ' [Computer program].', period(esc(item.publisher)), loc];
        break;
      }
    }
    return join(out, ' ');
  };

  /* ---- Vancouver ---- */
  FORMATTERS.vancouver = function (item) {
    var A = authorsVancouver(item.authors), E = item.editors || [];
    var y = yearOf(item);
    var loc = locator(item, 'vancouver');
    var t = period(esc(item.title));
    var out = [];
    switch (item.type) {
      case 'journalArticle': case 'other': {
        var src = esc(item.containerTitle) + '. ' + y + (has(item.month) ? ' ' + monthName(item.month, MONTHS_3) : '') + (has(item.volume) ? ';' + esc(item.volume) : '') + (has(item.issue) ? '(' + esc(item.issue) + ')' : '') + (has(item.pages) ? ':' + esc(hyphen(item.pages)) : '') + '.';
        out = [A ? period(A) : '', t, src, loc];
        break;
      }
      case 'book': {
        if (!A && E.length) A = authorsVancouver(E) + (E.length > 1 ? ', editors' : ', editor');
        out = [A ? period(A) : '', t, has(item.edition) ? ordinal(item.edition) + ' ed.' : '', join([esc(item.place), esc(item.publisher)], ': ') + '; ' + y + '.', loc];
        break;
      }
      case 'bookSection': {
        var eds = E.length ? 'In: ' + authorsVancouver(E) + (E.length > 1 ? ', editors. ' : ', editor. ') : 'In: ';
        out = [A ? period(A) : '', t, eds + period(esc(item.containerTitle)), join([esc(item.place), esc(item.publisher)], ': ') + '; ' + y + '.', has(item.pages) ? 'p. ' + esc(hyphen(item.pages)) + '.' : '', loc];
        break;
      }
      case 'conferencePaper': {
        out = [A ? period(A) : '', t, 'In: ' + esc(item.containerTitle) + '; ' + y + (has(item.place) ? '; ' + esc(item.place) : '') + '.', has(item.publisher) ? esc(item.publisher) + '; ' + y + '.' : '', has(item.pages) ? 'p. ' + esc(hyphen(item.pages)) + '.' : '', loc];
        break;
      }
      case 'thesis': {
        out = [A ? period(A) : '', esc(item.title) + ' [' + thesisGenre(item, 'vancouver') + '].', join([esc(item.place), esc(item.institution)], ': ') + '; ' + y + '.', loc];
        break;
      }
      case 'report': {
        out = [A ? period(A) : '', t, join([esc(item.place), esc(item.institution || item.publisher)], ': ') + '; ' + y + '.', has(item.number) ? 'Report No.: ' + esc(item.number) + '.' : '', loc];
        break;
      }
      case 'webpage': {
        var cited = has(item.accessed) ? ' [cited ' + fmtAccessed(item.accessed, 'vancouver') + ']' : '';
        out = [A ? period(A) : '', esc(item.title) + ' [Internet].', join([esc(item.place), esc(item.containerTitle)], ': ') + '; ' + y + cited + '.', has(item.url) ? 'Available from: ' + esc(item.url) : ''];
        break;
      }
      case 'preprint': {
        out = [A ? period(A) : '', t, join([esc(item.containerTitle), esc(item.number)], ' ') + ' [Preprint]. ' + y + '.', loc];
        break;
      }
      case 'dataset': {
        out = [A ? period(A) : '', esc(item.title) + ' [dataset].', esc(item.publisher) + '; ' + y + '.', loc];
        break;
      }
      case 'patent': {
        out = [A ? period(A) : '', esc(item.title) + '.', esc(join([item.place, 'patent ' + item.number], ' ')) + '. ' + y + '.', loc];
        break;
      }
      case 'standard': {
        out = [esc(item.publisher || item.institution || A) + '.', t, esc(item.number) + '. ' + join([esc(item.place), esc(item.publisher)], ': ') + '; ' + y + '.', loc];
        break;
      }
      case 'software': {
        out = [A ? period(A) : '', esc(item.title) + ' [software].', join([has(item.version) ? 'Version ' + esc(item.version) : '', esc(item.publisher)], '. ') + '; ' + y + '.', loc];
        break;
      }
    }
    return join(out, ' ');
  };

  function formatReference(item, style, index, opts) {
    style = STYLES[style] ? style : 'apa';
    opts = opts || {};
    item = normalizeItem(item);
    var html = FORMATTERS[style](item);
    html = html.replace(/\s+/g, ' ').replace(/([^.\s])\.\.(?![.])/g, '$1.').replace(/,\s*\.(?!\s*\.)/g, '.').trim();
    if (opts.markRetracted && item.retracted) html = 'RETRACTED: ' + html;
    var prefix = '';
    if (STYLES[style].numeric && index != null) prefix = style === 'ieee' ? '[' + index + '] ' : index + '. ';
    return { html: prefix + html, text: stripHtml(prefix + html), prefix: prefix };
  }

  /* ------------------------------------------------------------------ */
  /*  In-text citations                                                  */
  /* ------------------------------------------------------------------ */

  function citeAuthors(item, style) {
    var names = item.authors && item.authors.length ? item.authors : (item.editors || []);
    if (!names.length) {
      // No author: use short title
      var t = trim(item.title) || 'Untitled';
      var short = t.split(/\s+/).slice(0, 4).join(' ');
      return (item.type === 'book' || item.type === 'report' || item.type === 'webpage' ? '<i>' + esc(short) + '</i>' : '\u201C' + esc(short) + '\u201D');
    }
    var f = family;
    switch (style) {
      case 'apa':
        if (names.length === 1) return esc(f(names[0]));
        if (names.length === 2) return esc(f(names[0])) + ' & ' + esc(f(names[1]));
        return esc(f(names[0])) + ' et al.';
      case 'mla':
        if (names.length === 1) return esc(f(names[0]));
        if (names.length === 2) return esc(f(names[0])) + ' and ' + esc(f(names[1]));
        return esc(f(names[0])) + ' et al.';
      case 'chicago':
        if (names.length === 1) return esc(f(names[0]));
        if (names.length === 2) return esc(f(names[0])) + ' and ' + esc(f(names[1]));
        if (names.length === 3) return esc(f(names[0])) + ', ' + esc(f(names[1])) + ', and ' + esc(f(names[2]));
        return esc(f(names[0])) + ' et al.';
      case 'harvard':
        if (names.length === 1) return esc(f(names[0]));
        if (names.length === 2) return esc(f(names[0])) + ' and ' + esc(f(names[1]));
        if (names.length === 3) return esc(f(names[0])) + ', ' + esc(f(names[1])) + ' and ' + esc(f(names[2]));
        return esc(f(names[0])) + ' et al.';
    }
    return esc(f(names[0]));
  }

  // items: array of items (or one item); opts: { numbers: [n,...] for numeric styles, pages: 'p. 12', narrative: bool }
  function formatCitation(items, style, opts) {
    style = STYLES[style] ? style : 'apa';
    opts = opts || {};
    if (!Array.isArray(items)) items = [items];
    items = items.map(normalizeItem);
    var html;
    if (STYLES[style].numeric) {
      var nums = (opts.numbers || items.map(function (_, i) { return i + 1; })).slice().sort(function (a, b) { return a - b; });
      // compress ranges: 1,2,3,5 -> 1–3, 5
      var groups = [], start = null, prev = null;
      nums.forEach(function (n) {
        if (start === null) { start = prev = n; return; }
        if (n === prev + 1) { prev = n; return; }
        groups.push(start === prev ? String(start) : (prev === start + 1 ? start + ', ' + prev : start + '\u2013' + prev));
        start = prev = n;
      });
      if (start !== null) groups.push(start === prev ? String(start) : (prev === start + 1 ? start + ', ' + prev : start + '\u2013' + prev));
      var inner = groups.join(', ');
      if (opts.pages) inner += ', ' + esc(opts.pages);
      if (style === 'vancouver') html = opts.superscript ? '<sup>' + inner + '</sup>' : '(' + inner + ')';
      else html = '[' + inner + ']';
      return { html: html, text: stripHtml(html) };
    }
    var parts = items.map(function (item) {
      var a = citeAuthors(item, style), y = yearOf(item), pg = opts.pages ? esc(opts.pages) : '';
      switch (style) {
        case 'apa':
          return opts.narrative ? a + ' (' + y + (pg ? ', ' + pg : '') + ')' : a + ', ' + y + (pg ? ', ' + pg : '');
        case 'mla':
          return opts.narrative ? a + (pg ? ' (' + pg + ')' : '') : a + (pg ? ' ' + pg : '');
        case 'chicago':
          return opts.narrative ? a + ' (' + y + (pg ? ', ' + pg : '') + ')' : a + ' ' + y + (pg ? ', ' + pg : '');
        case 'harvard':
          return opts.narrative ? a + ' (' + y + (pg ? ', ' + pg : '') + ')' : a + ', ' + y + (pg ? ', ' + pg : '');
      }
      return a + ', ' + y;
    });
    if (opts.narrative) html = parts.join('; ');
    else html = '(' + parts.join('; ') + ')';
    return { html: html, text: stripHtml(html) };
  }

  /* ------------------------------------------------------------------ */
  /*  Bibliography                                                       */
  /* ------------------------------------------------------------------ */

  function sortKey(item) {
    var names = item.authors && item.authors.length ? item.authors : (item.editors || []);
    var a = names.length ? names.map(function (n) { return (family(n) + ' ' + (n.given || '')).toLowerCase(); }).join(' ') : trim(item.title).toLowerCase().replace(/^(the|a|an)\s+/, '');
    return a + '\u0000' + (has(item.year) ? item.year : 9999) + '\u0000' + trim(item.title).toLowerCase();
  }
  function sortForBibliography(items, style) {
    if (STYLES[style] && STYLES[style].numeric) return items.slice();
    return items.slice().sort(function (a, b) { var ka = sortKey(a), kb = sortKey(b); return ka < kb ? -1 : ka > kb ? 1 : 0; });
  }
  function bibliography(items, style, opts) {
    opts = opts || {};
    style = STYLES[style] ? style : 'apa';
    var sorted = opts.keepOrder ? items.slice() : sortForBibliography(items, style);
    var entries = sorted.map(function (item, i) {
      var r = formatReference(item, style, STYLES[style].numeric ? i + 1 : null, opts);
      return { item: item, n: i + 1, html: r.html, text: r.text };
    });
    var hanging = !STYLES[style].numeric;
    var pStyle = hanging ? 'margin:0 0 0 0.5in;text-indent:-0.5in;line-height:2' : 'margin:0 0 6pt 0;line-height:1.5';
    var html = entries.map(function (e) { return '<p style="' + pStyle + '">' + e.html + '</p>'; }).join('\n');
    return { entries: entries, html: html, text: entries.map(function (e) { return e.text; }).join('\n') };
  }

  /* ------------------------------------------------------------------ */
  /*  Duplicates & cite keys                                             */
  /* ------------------------------------------------------------------ */

  function normTitle(t) { return trim(t).toLowerCase().replace(/<[^>]+>/g, '').replace(/[^a-z0-9\u0600-\u06FF]+/g, ' ').trim(); }
  function findDuplicates(items) {
    var byDoi = {}, byTitle = {}, groups = [], seen = {};
    items.forEach(function (it) {
      var d = cleanDoi(it.doi).toLowerCase();
      if (d) (byDoi[d] = byDoi[d] || []).push(it);
      var t = normTitle(it.title);
      if (t.length > 12) (byTitle[t] = byTitle[t] || []).push(it);
    });
    function add(list) {
      var ids = list.map(function (x) { return x.id; }).sort().join('|');
      if (list.length > 1 && !seen[ids]) { seen[ids] = true; groups.push(list); }
    }
    Object.keys(byDoi).forEach(function (k) { add(byDoi[k]); });
    Object.keys(byTitle).forEach(function (k) { add(byTitle[k]); });
    return groups;
  }
  function citeKey(item, existing) {
    var names = item.authors && item.authors.length ? item.authors : item.editors || [];
    var fam = names.length ? family(names[0]) : 'anon';
    fam = fam.normalize ? fam.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : fam;
    fam = fam.toLowerCase().replace(/[^a-z0-9]/g, '') || 'anon';
    var words = trim(item.title).toLowerCase().replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(function (w) { return w && !/^(a|an|the|on|of|in|for|and|to|with|from|by|at|is|are|we|our|its)$/.test(w); });
    var key = fam + (has(item.year) ? item.year : '') + (words[0] || '');
    if (existing) { var base = key, i = 2; while (existing[key]) key = base + String.fromCharCode(95 + i++); }
    return key;
  }

  /* ------------------------------------------------------------------ */
  /*  BibTeX                                                             */
  /* ------------------------------------------------------------------ */

  var LATEX_MAP = {
    "\\'a": 'á', "\\'e": 'é', "\\'i": 'í', "\\'o": 'ó', "\\'u": 'ú', "\\'y": 'ý', "\\'A": 'Á', "\\'E": 'É', "\\'I": 'Í', "\\'O": 'Ó', "\\'U": 'Ú', "\\'c": 'ć', "\\'n": 'ń', "\\'s": 'ś', "\\'z": 'ź',
    '\\`a': 'à', '\\`e': 'è', '\\`i': 'ì', '\\`o': 'ò', '\\`u': 'ù', '\\`A': 'À', '\\`E': 'È',
    '\\"a': 'ä', '\\"e': 'ë', '\\"i': 'ï', '\\"o': 'ö', '\\"u': 'ü', '\\"A': 'Ä', '\\"O': 'Ö', '\\"U': 'Ü', '\\"y': 'ÿ',
    '\\^a': 'â', '\\^e': 'ê', '\\^i': 'î', '\\^o': 'ô', '\\^u': 'û', '\\~n': 'ñ', '\\~a': 'ã', '\\~o': 'õ', '\\~N': 'Ñ',
    '\\c{c}': 'ç', '\\c{C}': 'Ç', '\\v{c}': 'č', '\\v{s}': 'š', '\\v{z}': 'ž', '\\v{C}': 'Č', '\\v{S}': 'Š', '\\v{Z}': 'Ž', '\\v{r}': 'ř', '\\v{e}': 'ě',
    '\\ss': 'ß', '\\o': 'ø', '\\O': 'Ø', '\\aa': 'å', '\\AA': 'Å', '\\ae': 'æ', '\\AE': 'Æ', '\\l': 'ł', '\\L': 'Ł', '\\i': 'ı',
    '\\&': '&', '\\%': '%', '\\$': '$', '\\#': '#', '\\_': '_', '\\{': '{', '\\}': '}', '---': '—', '--': '–', '~': ' ', "''": '\u201D', '``': '\u201C', "\\textendash": '–', '\\textemdash': '—'
  };
  function delatex(s) {
    s = String(s || '');
    // \'{e} -> \'e
    s = s.replace(/\\(['`"^~])\{(\w)\}/g, '\\$1$2');
    s = s.replace(/\{\\(['`"^~])(\w)\}/g, '\\$1$2');
    Object.keys(LATEX_MAP).sort(function (a, b) { return b.length - a.length; }).forEach(function (k) { s = s.split(k).join(LATEX_MAP[k]); });
    s = s.replace(/\\(emph|textit|textbf|textsc|mathrm|url)\{([^{}]*)\}/g, '$2');
    s = s.replace(/[{}]/g, '').replace(/\\\s/g, ' ').replace(/\s+/g, ' ').trim();
    return s;
  }
  function latexEscape(s) { return String(s || '').replace(/([&%$#_])/g, '\\$1'); }
  // Split a BibTeX author/editor field on " and " at brace depth 0; {Braced Names} become literal (organization) names.
  function bibNames(raw) {
    var s = String(raw || ''), parts = [], depth = 0, cur = '';
    for (var i = 0; i < s.length; i++) {
      var ch = s[i];
      if (ch === '{') depth++; else if (ch === '}') depth--;
      if (depth === 0 && /\s/.test(ch)) {
        var m = s.slice(i).match(/^\s+and\s+/i);
        if (m) { parts.push(cur); cur = ''; i += m[0].length - 1; continue; }
      }
      cur += ch;
    }
    if (trim(cur)) parts.push(cur);
    var out = [];
    parts.forEach(function (p) {
      p = trim(p); if (!p) return;
      if (/^\{.*\}$/.test(p) && p.indexOf(',') < 0) { out.push({ literal: delatex(p.slice(1, -1)) }); return; }
      var n = parseNames(delatex(p).replace(/\s+and\s+/gi, ' '));
      if (n[0]) out.push(n[0]);
    });
    return out;
  }
  function bibNameField(names) {
    return names.map(function (n) { return n.literal ? '{' + latexEscape(n.literal) + '}' : latexEscape(n.family) + (n.given ? ', ' + latexEscape(n.given) : ''); }).join(' and ');
  }

  var BIB_TYPE_IN = { article: 'journalArticle', book: 'book', inbook: 'bookSection', incollection: 'bookSection', inproceedings: 'conferencePaper', conference: 'conferencePaper', proceedings: 'book', phdthesis: 'thesis', mastersthesis: 'thesis', thesis: 'thesis', techreport: 'report', report: 'report', misc: 'other', online: 'webpage', electronic: 'webpage', www: 'webpage', patent: 'patent', dataset: 'dataset', software: 'software', standard: 'standard', unpublished: 'other', manual: 'report', booklet: 'book' };
  var BIB_TYPE_OUT = { journalArticle: 'article', book: 'book', bookSection: 'incollection', conferencePaper: 'inproceedings', thesis: 'phdthesis', report: 'techreport', webpage: 'online', preprint: 'misc', dataset: 'dataset', patent: 'patent', standard: 'standard', software: 'software', other: 'misc' };

  function bibtexParse(text) {
    var items = [], i = 0, s = String(text || '');
    while (i < s.length) {
      var at = s.indexOf('@', i); if (at < 0) break;
      var m = s.slice(at).match(/^@(\w+)\s*[{(]/); if (!m) { i = at + 1; continue; }
      var type = m[1].toLowerCase(); i = at + m[0].length;
      if (type === 'comment' || type === 'preamble') { i = s.indexOf('\n', i); if (i < 0) break; continue; }
      if (type === 'string') { var e = s.indexOf('}', i); i = e < 0 ? s.length : e + 1; continue; }
      // read key
      var comma = s.indexOf(',', i), close = s.indexOf('}', i);
      var key = trim(s.slice(i, comma < 0 || (close >= 0 && close < comma) ? close : comma));
      i = comma < 0 ? i : comma + 1;
      var fields = {}, depth = 1;
      // parse fields until the matching close brace
      while (i < s.length) {
        while (i < s.length && /[\s,]/.test(s[i])) i++;
        if (s[i] === '}' || s[i] === ')') { i++; break; }
        var fm = s.slice(i).match(/^([\w\-:+.]+)\s*=\s*/); if (!fm) { i++; continue; }
        var name = fm[1].toLowerCase(); i += fm[0].length;
        var val = '';
        // concatenated values with #
        while (true) {
          while (i < s.length && /\s/.test(s[i])) i++;
          if (s[i] === '{') {
            var d = 0, start = i;
            for (; i < s.length; i++) { if (s[i] === '{') d++; else if (s[i] === '}') { d--; if (d === 0) { i++; break; } } }
            val += s.slice(start + 1, i - 1);
          } else if (s[i] === '"') {
            var st = ++i, dd = 0;
            for (; i < s.length; i++) { if (s[i] === '{') dd++; else if (s[i] === '}') dd--; else if (s[i] === '"' && dd === 0) break; }
            val += s.slice(st, i); i++;
          } else {
            var nm = s.slice(i).match(/^[^,}\s#]+/); if (nm) { val += nm[0]; i += nm[0].length; } else i++;
          }
          while (i < s.length && /\s/.test(s[i])) i++;
          if (s[i] === '#') { i++; continue; }
          break;
        }
        fields[name] = val;
      }
      items.push(bibToItem(type, key, fields));
    }
    return items;
  }
  function bibToItem(type, key, f) {
    var it = newItem(BIB_TYPE_IN[type] || 'other');
    if (type === 'misc' || type === 'unpublished') {
      if (f.eprint || f.archiveprefix || /arxiv/i.test(f.journal || '') || /arxiv/i.test(f.howpublished || '')) it.type = 'preprint';
      else if (f.url || /^\\?url/.test(f.howpublished || '') || f.howpublished && /http/.test(f.howpublished)) it.type = 'webpage';
    }
    it.citeKey = key;
    it.title = delatex(f.title);
    it.authors = bibNames(f.author);
    it.editors = bibNames(f.editor);
    if (f.year) { var y = String(f.year).match(/\d{4}/); it.year = y ? +y[0] : ''; }
    if (f.month) { var mo = String(f.month).toLowerCase().slice(0, 3), idx = MONTHS_3.map(function (x) { return x.toLowerCase(); }).indexOf(mo); it.month = idx >= 0 ? idx + 1 : (parseInt(f.month, 10) || ''); }
    if (f.date && !f.year) { var pd = parseDateStr(f.date); if (pd) { it.year = pd.y; it.month = pd.m || ''; it.day = pd.d || ''; } }
    it.containerTitle = delatex(f.journal || f.journaltitle || f.booktitle || '');
    if (it.type === 'webpage') it.containerTitle = delatex(f.organization || f.publisher || (f.howpublished && !/http/.test(f.howpublished) ? f.howpublished : '') || '');
    if (it.type === 'preprint') { it.containerTitle = delatex(f.archiveprefix || f.eprinttype || f.journal || f.howpublished || 'arXiv'); it.number = delatex(f.eprint ? (/arxiv/i.test(f.archiveprefix || f.eprinttype || 'arXiv') ? 'arXiv:' + String(f.eprint).replace(/^arxiv:\s*/i, '') : f.eprint) : f.number || ''); }
    it.volume = delatex(f.volume);
    if (it.type === 'journalArticle') it.issue = delatex(f.number || f.issue);
    else { it.issue = delatex(f.issue); if (it.type !== 'preprint') it.number = delatex(f.number); }
    it.pages = delatex(f.pages); it.publisher = delatex(f.publisher || (it.type === 'thesis' ? '' : f.organization || '')); it.place = delatex(f.address || f.location);
    it.edition = delatex(f.edition); it.series = it.type === 'book' ? delatex(f.series) : ''; it.institution = delatex(f.school || f.institution || f.organization);
    if (it.type === 'thesis') it.genre = type === 'mastersthesis' ? "Master's thesis" : (delatex(f.type) || 'PhD thesis');
    else it.genre = delatex(f.type);
    it.version = delatex(f.version); it.isbn = delatex(f.isbn); it.issn = delatex(f.issn); it.doi = cleanDoi(delatex(f.doi));
    it.url = delatex(f.url || (f.howpublished && /https?:/.test(f.howpublished) ? f.howpublished.replace(/\\url\{([^}]*)\}/, '$1') : ''));
    it.accessed = delatex(f.urldate); it.language = delatex(f.language); it.abstract = delatex(f.abstract);
    it.tags = (delatex(f.keywords) ? delatex(f.keywords).split(/[,;]/).map(trim).filter(Boolean) : []);
    it.notes = delatex(f.note || f.annote || '');
    if (f.howpublished && it.type === 'other') it.containerTitle = delatex(f.howpublished);
    return normalizeItem(it);
  }
  function bibtexSerialize(items) {
    var keys = {};
    return items.map(function (item) {
      item = normalizeItem(item);
      var key = item.citeKey || citeKey(item, keys); keys[key] = true;
      var t = BIB_TYPE_OUT[item.type] || 'misc';
      if (item.type === 'thesis' && /master/i.test(item.genre)) t = 'mastersthesis';
      var f = [];
      function add(k, v) { if (has(v)) f.push('  ' + k + ' = {' + latexEscape(v) + '}'); }
      add('title', item.title);
      if (item.authors.length) f.push('  author = {' + bibNameField(item.authors) + '}');
      if (item.editors.length) f.push('  editor = {' + bibNameField(item.editors) + '}');
      add('year', item.year); if (has(item.month)) add('month', MONTHS_3[+item.month - 1] ? MONTHS_3[+item.month - 1].toLowerCase() : item.month);
      if (item.type === 'journalArticle') { add('journal', item.containerTitle); add('volume', item.volume); add('number', item.issue); }
      else if (item.type === 'bookSection' || item.type === 'conferencePaper') { add('booktitle', item.containerTitle); add('volume', item.volume); }
      else if (item.type === 'webpage') { add('organization', item.containerTitle); }
      else if (item.type === 'preprint') { add('eprinttype', 'arxiv'); add('eprint', String(item.number || '').replace(/^arxiv:\s*/i, '')); add('journal', item.containerTitle); }
      else if (item.type === 'book') { add('series', item.series); add('volume', item.volume); }
      else add('howpublished', item.containerTitle);
      add('pages', enDash(item.pages).replace(/\u2013/g, '--')); add('publisher', item.publisher); add('address', item.place); add('edition', item.edition);
      if (item.type === 'thesis') { add('school', item.institution); add('type', item.genre); }
      else if (item.type === 'report') { add('institution', item.institution); add('number', item.number); add('type', item.genre); }
      else if (item.type !== 'journalArticle' && item.type !== 'preprint') { add('organization', item.institution); add('number', item.number); }
      add('version', item.version); add('isbn', item.isbn); add('issn', item.issn); add('doi', item.doi); add('url', item.url); add('urldate', item.accessed);
      add('language', item.language); add('abstract', item.abstract); add('keywords', item.tags.join(', ')); add('note', item.notes);
      return '@' + t + '{' + key + ',\n' + f.join(',\n') + '\n}';
    }).join('\n\n') + '\n';
  }

  /* ------------------------------------------------------------------ */
  /*  RIS                                                                */
  /* ------------------------------------------------------------------ */

  var RIS_TYPE_IN = { JOUR: 'journalArticle', EJOUR: 'journalArticle', MGZN: 'journalArticle', NEWS: 'journalArticle', BOOK: 'book', EBOOK: 'book', EDBOOK: 'book', CHAP: 'bookSection', ECHAP: 'bookSection', CONF: 'conferencePaper', CPAPER: 'conferencePaper', THES: 'thesis', RPRT: 'report', ELEC: 'webpage', WEB: 'webpage', ICOMM: 'webpage', BLOG: 'webpage', DATA: 'dataset', DBASE: 'dataset', PAT: 'patent', STAND: 'standard', COMP: 'software', UNPB: 'preprint', MANSCPT: 'other', GEN: 'other' };
  var RIS_TYPE_OUT = { journalArticle: 'JOUR', book: 'BOOK', bookSection: 'CHAP', conferencePaper: 'CONF', thesis: 'THES', report: 'RPRT', webpage: 'ELEC', preprint: 'UNPB', dataset: 'DATA', patent: 'PAT', standard: 'STAND', software: 'COMP', other: 'GEN' };

  function risParse(text) {
    var items = [], cur = null, lines = String(text || '').replace(/\r\n?/g, '\n').split('\n');
    function finish() { if (cur) { items.push(risToItem(cur)); cur = null; } }
    lines.forEach(function (line) {
      var m = line.match(/^([A-Z][A-Z0-9])\s{1,2}-\s?(.*)$/);
      if (!m) { if (cur && cur._last && trim(line)) cur[cur._last][cur[cur._last].length - 1] += ' ' + trim(line); return; }
      var tag = m[1], val = trim(m[2]);
      if (tag === 'TY') { finish(); cur = { TY: [val] }; cur._last = 'TY'; return; }
      if (!cur) return;
      if (tag === 'ER') { finish(); return; }
      (cur[tag] = cur[tag] || []).push(val); cur._last = tag;
    });
    finish();
    return items;
  }
  function risToItem(r) {
    function g(t) { return r[t] ? r[t][0] : ''; }
    function all(t) { return r[t] || []; }
    var it = newItem(RIS_TYPE_IN[g('TY')] || 'other');
    it.title = g('TI') || g('T1') || g('CT');
    it.authors = normNames(all('AU').concat(all('A1')));
    it.editors = normNames(all('ED').concat(all('A2').filter(function () { return it.type === 'bookSection' || it.type === 'book'; })));
    var py = g('PY') || g('Y1') || g('DA');
    if (py) { var pd = py.match(/^(\d{4})(?:\/(\d{0,2})(?:\/(\d{0,2}))?)?/); if (pd) { it.year = +pd[1]; it.month = pd[2] ? +pd[2] : ''; it.day = pd[3] ? +pd[3] : ''; } else { var d2 = parseDateStr(py); if (d2) { it.year = d2.y; it.month = d2.m || ''; it.day = d2.d || ''; } } }
    it.containerTitle = g('T2') || g('JO') || g('JF') || g('J2') || g('BT') || '';
    if (it.type === 'journalArticle') it.containerTitle = g('JF') || g('JO') || g('T2') || g('J2') || '';
    it.volume = g('VL'); it.issue = g('IS');
    var sp = g('SP'), ep = g('EP'); it.pages = sp ? (ep && ep !== sp ? sp + '–' + ep : sp) : '';
    it.publisher = g('PB'); it.place = g('CY'); it.edition = g('ET'); it.series = g('T3');
    it.institution = it.type === 'thesis' || it.type === 'report' || it.type === 'patent' || it.type === 'standard' ? (g('PB') || g('A3') || g('C1') || '') : '';
    if (it.type === 'thesis' || it.type === 'report') it.publisher = '';
    it.genre = g('M3'); it.number = g('M1') || (it.type === 'standard' ? g('SN') : '');
    it.version = g('ET') && it.type === 'software' ? g('ET') : ''; if (it.type === 'software') it.edition = '';
    var sn = g('SN'); if (it.type === 'journalArticle') it.issn = sn; else if (it.type === 'book' || it.type === 'bookSection') it.isbn = sn;
    it.doi = cleanDoi(g('DO') || (all('UR').concat(all('L1')).filter(function (u) { return /doi\.org/.test(u); })[0] || ''));
    it.url = all('UR').filter(function (u) { return !/doi\.org/.test(u); })[0] || g('L2') || '';
    it.accessed = g('Y2'); it.language = g('LA'); it.abstract = g('AB') || g('N2');
    it.tags = all('KW').map(trim).filter(Boolean); it.notes = all('N1').join('\n');
    if (it.type === 'preprint') it.number = g('M1') || g('SN');
    return normalizeItem(it);
  }
  function risSerialize(items) {
    return items.map(function (item) {
      item = normalizeItem(item);
      var L = [];
      function add(t, v) { if (has(v)) L.push(t + '  - ' + v); }
      add('TY', RIS_TYPE_OUT[item.type] || 'GEN');
      add('TI', item.title);
      item.authors.forEach(function (n) { add('AU', n.literal ? n.literal : n.family + (n.given ? ', ' + n.given : '')); });
      item.editors.forEach(function (n) { add('ED', n.literal ? n.literal : n.family + (n.given ? ', ' + n.given : '')); });
      if (has(item.year)) add('PY', item.year + (has(item.month) ? '/' + pad2(+item.month) + (has(item.day) ? '/' + pad2(+item.day) : '/') : '//'));
      add(item.type === 'journalArticle' ? 'JF' : 'T2', item.containerTitle);
      add('VL', item.volume); add('IS', item.issue);
      if (has(item.pages)) { var p = enDash(item.pages).split('–'); add('SP', p[0]); if (p[1]) add('EP', p[1]); }
      add('PB', item.publisher || item.institution); add('CY', item.place); add('ET', item.edition || item.version); add('T3', item.series);
      add('M3', item.genre); add('M1', item.number); add('SN', item.isbn || item.issn);
      add('DO', item.doi); add('UR', item.url); add('Y2', item.accessed); add('LA', item.language); add('AB', item.abstract);
      item.tags.forEach(function (t) { add('KW', t); });
      if (has(item.notes)) add('N1', item.notes.replace(/\s*\n\s*/g, ' '));
      L.push('ER  - ');
      return L.join('\n');
    }).join('\n\n') + '\n';
  }

  /* ------------------------------------------------------------------ */
  /*  CSL-JSON (also used for Crossref)                                  */
  /* ------------------------------------------------------------------ */

  var CSL_TYPE_IN = { 'article-journal': 'journalArticle', 'article-magazine': 'journalArticle', 'article-newspaper': 'journalArticle', 'journal-article': 'journalArticle', article: 'preprint', 'posted-content': 'preprint', book: 'book', monograph: 'book', 'edited-book': 'book', 'reference-book': 'book', chapter: 'bookSection', 'book-chapter': 'bookSection', 'book-section': 'bookSection', 'book-part': 'bookSection', 'paper-conference': 'conferencePaper', 'proceedings-article': 'conferencePaper', proceedings: 'book', thesis: 'thesis', dissertation: 'thesis', report: 'report', webpage: 'webpage', 'post-weblog': 'webpage', post: 'webpage', dataset: 'dataset', patent: 'patent', standard: 'standard', software: 'software', 'computer-program': 'software', 'peer-review': 'other', 'journal-issue': 'other' };
  var CSL_TYPE_OUT = { journalArticle: 'article-journal', book: 'book', bookSection: 'chapter', conferencePaper: 'paper-conference', thesis: 'thesis', report: 'report', webpage: 'webpage', preprint: 'article', dataset: 'dataset', patent: 'patent', standard: 'standard', software: 'software', other: 'document' };

  function firstOf(v) { return Array.isArray(v) ? (v.length ? v[0] : '') : (v == null ? '' : v); }
  function stripJats(s) { return String(s || '').replace(/<\/?jats:[^>]+>/g, '').replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim(); }
  function fromCSL(c) {
    var it = newItem(CSL_TYPE_IN[c.type] || 'other');
    it.title = stripJats(firstOf(c.title));
    if (c.subtitle && firstOf(c.subtitle) && it.title.indexOf(firstOf(c.subtitle)) < 0) it.title += ': ' + firstOf(c.subtitle);
    it.authors = normNames((c.author || []).map(function (a) { return a.literal || (!a.family && a.name) ? { literal: a.literal || a.name } : { family: a.family || '', given: a.given || '' }; }));
    it.editors = normNames((c.editor || []).map(function (a) { return a.literal ? { literal: a.literal } : { family: a.family || '', given: a.given || '' }; }));
    var issued = c.issued && c.issued['date-parts'] && c.issued['date-parts'][0] || (c['published-print'] || c['published-online'] || c.created || {})['date-parts'] && (c['published-print'] || c['published-online'] || c.created)['date-parts'][0];
    if (issued) { it.year = issued[0] || ''; it.month = issued[1] || ''; it.day = issued[2] || ''; }
    it.containerTitle = stripJats(firstOf(c['container-title']) || firstOf(c['collection-title']) || (c.event && (c.event.name || c.event)) || '');
    if (it.type === 'conferencePaper' && c.event && c.event.name && !firstOf(c['container-title'])) it.containerTitle = c.event.name;
    it.volume = firstOf(c.volume); it.issue = firstOf(c.issue); it.pages = enDash(firstOf(c.page) || (has(c['page-first']) ? c['page-first'] : ''));
    it.publisher = firstOf(c.publisher); it.place = firstOf(c['publisher-place']) || firstOf(c['publisher-location']) || firstOf(c['event-place']) || '';
    it.edition = firstOf(c.edition); it.series = firstOf(c['collection-title']) && it.type === 'book' ? firstOf(c['collection-title']) : '';
    it.institution = it.type === 'thesis' || it.type === 'report' || it.type === 'standard' || it.type === 'patent' ? (firstOf(c.publisher) || firstOf(c.institution) || '') : '';
    if (it.type === 'thesis' || it.type === 'report') it.publisher = '';
    it.genre = firstOf(c.genre); it.number = firstOf(c.number) || firstOf(c['collection-number']) || '';
    it.version = firstOf(c.version); it.isbn = firstOf(c.ISBN); it.issn = firstOf(c.ISSN);
    it.doi = cleanDoi(firstOf(c.DOI)); it.url = cleanDoi(firstOf(c.DOI)) ? (firstOf(c.URL) && !/doi\.org/.test(firstOf(c.URL)) ? firstOf(c.URL) : '') : firstOf(c.URL) || (c.resource && c.resource.primary && c.resource.primary.URL) || '';
    if (c.accessed && c.accessed['date-parts'] && c.accessed['date-parts'][0]) { var a = c.accessed['date-parts'][0]; it.accessed = a[0] + (a[1] ? '-' + pad2(a[1]) : '') + (a[2] ? '-' + pad2(a[2]) : ''); }
    it.language = firstOf(c.language); it.abstract = stripJats(firstOf(c.abstract));
    it.tags = (c.keyword ? String(c.keyword).split(/[,;]/) : (c.subject || [])).map(trim).filter(Boolean).slice(0, 10);
    it.notes = firstOf(c.note) || '';
    // Crossref retraction flags
    var upd = c['updated-by'] || [];
    for (var i = 0; i < upd.length; i++) if (/retract|withdraw/i.test(upd[i].type || upd[i].label || '')) { it.retracted = true; it.retractionNote = (upd[i].label || 'Retraction') + (upd[i].DOI ? ' — notice DOI ' + upd[i].DOI : ''); }
    if (c['is-retracted'] || c.is_retracted) it.retracted = true;
    if (/retract/i.test(it.title) && /^(retracted|withdrawn)[:\s]/i.test(it.title)) it.retracted = true;
    if (it.type === 'preprint' && /arxiv/i.test(it.doi + ' ' + it.url + ' ' + it.publisher)) { it.containerTitle = it.containerTitle || 'arXiv'; var am = (it.doi + ' ' + it.url).match(/(\d{4}\.\d{4,5}(v\d+)?)/); if (am && !it.number) it.number = 'arXiv:' + am[1]; }
    return normalizeItem(it);
  }
  function toCSL(item) {
    item = normalizeItem(item);
    var c = { id: item.id, type: CSL_TYPE_OUT[item.type] || 'document', title: item.title };
    function names(list) { return list.map(function (n) { return n.literal ? { literal: n.literal } : { family: n.family, given: n.given }; }); }
    if (item.authors.length) c.author = names(item.authors);
    if (item.editors.length) c.editor = names(item.editors);
    if (has(item.year)) { var dp = [+item.year]; if (has(item.month)) { dp.push(+item.month); if (has(item.day)) dp.push(+item.day); } c.issued = { 'date-parts': [dp] }; }
    function set(k, v) { if (has(v)) c[k] = v; }
    set('container-title', item.containerTitle); set('volume', item.volume); set('issue', item.issue); set('page', item.pages);
    set('publisher', item.publisher || item.institution); set('publisher-place', item.place); set('edition', item.edition); set('collection-title', item.series);
    set('genre', item.genre); set('number', item.number); set('version', item.version); set('ISBN', item.isbn); set('ISSN', item.issn); set('DOI', item.doi); set('URL', item.url);
    if (has(item.accessed)) { var ad = parseDateStr(item.accessed); if (ad) c.accessed = { 'date-parts': [[ad.y].concat(ad.m ? [ad.m].concat(ad.d ? [ad.d] : []) : [])] }; }
    set('language', item.language); set('abstract', item.abstract); if (item.tags.length) c.keyword = item.tags.join(', '); set('note', item.notes);
    return c;
  }
  function cslParse(text) {
    var data = typeof text === 'string' ? JSON.parse(text) : text;
    if (data && data.format === 'lexiref-library') return libraryParse(data).items;
    if (data && data.message) data = data.message.items || data.message; // Crossref envelope
    if (!Array.isArray(data)) data = [data];
    return data.map(function (c) { return c && (Array.isArray(c.authors) || c.added || c.format) ? normalizeItem(c) : fromCSL(c || {}); });
  }

  /* ------------------------------------------------------------------ */
  /*  LexiRef library file                                               */
  /* ------------------------------------------------------------------ */

  function librarySerialize(lib) {
    return JSON.stringify({ format: 'lexiref-library', version: 1, exported: new Date().toISOString(), app: 'LexiRef', collections: lib.collections || [], items: (lib.items || []).map(normalizeItem) }, null, 2);
  }
  function libraryParse(dataOrText) {
    var data = typeof dataOrText === 'string' ? JSON.parse(dataOrText) : dataOrText;
    if (Array.isArray(data)) return { collections: [], items: cslParse(data) };
    if (data && data.format === 'lexiref-library') return { collections: data.collections || [], items: (data.items || []).map(normalizeItem) };
    if (data && data.items && Array.isArray(data.items)) return { collections: data.collections || [], items: data.items.map(normalizeItem) };
    return { collections: [], items: cslParse(data) };
  }

  // Detect the format of pasted/imported text and parse it.
  function importAny(text, filename) {
    var t = trim(text), name = (filename || '').toLowerCase();
    if (!t) return [];
    if (/\.(bib|bibtex)$/.test(name) || /^@\w+\s*[{(]/.test(t)) return bibtexParse(t);
    if (/\.ris$/.test(name) || /^TY\s{1,2}-\s/m.test(t)) return risParse(t);
    if (/^[\[{]/.test(t)) return libraryParse(t).items;
    // bare DOI list?
    var dois = t.match(/10\.\d{4,9}\/[^\s"<>]+/g);
    if (dois) return dois.map(function (d) { var it = newItem('journalArticle'); it.doi = cleanDoi(d); it.title = '(DOI pending lookup) ' + d; return it; });
    return [];
  }

  /* ------------------------------------------------------------------ */
  /*  Network: Crossref / OpenAlex (optional, needs internet)            */
  /* ------------------------------------------------------------------ */

  var contactEmail = ''; // optional: joins the Crossref/OpenAlex "polite pool" (faster, fewer rate limits)
  function setContactEmail(e) { contactEmail = trim(e); }
  function fetchJSON(url, timeoutMs, retry) {
    if (typeof fetch !== 'function') return Promise.reject(new Error('fetch not available'));
    if (contactEmail) url += (url.indexOf('?') >= 0 ? '&' : '?') + 'mailto=' + encodeURIComponent(contactEmail);
    var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, timeoutMs || 15000) : null;
    return fetch(url, { headers: { Accept: 'application/json' }, signal: ctrl ? ctrl.signal : undefined })
      .then(function (r) {
        if (r.status === 429 && retry !== false) { // rate limited: wait and retry once
          if (timer) clearTimeout(timer);
          return new Promise(function (res) { setTimeout(res, 2000); }).then(function () { return fetchJSON(url.replace(/[?&]mailto=[^&]*/, function (m) { return m.charAt(0) === '?' ? '?' : ''; }).replace(/\?&/, '?'), timeoutMs, false); });
        }
        if (!r.ok) throw new Error(r.status === 404 ? 'not found' : 'HTTP ' + r.status); return r.json();
      })
      .then(function (j) { if (timer) clearTimeout(timer); return j; }, function (e) { if (timer) clearTimeout(timer); throw e; });
  }
  // DataCite (arXiv, Zenodo, datasets, software DOIs) -> item
  var DC_TYPE = { Preprint: 'preprint', Dataset: 'dataset', Software: 'software', JournalArticle: 'journalArticle', Journal: 'journalArticle', Book: 'book', BookChapter: 'bookSection', ConferencePaper: 'conferencePaper', ConferenceProceeding: 'conferencePaper', Dissertation: 'thesis', Report: 'report', Standard: 'standard', Text: 'other', Other: 'other' };
  function fromDataCite(a) {
    a = a || {};
    var types = a.types || {}, it = newItem(DC_TYPE[types.resourceTypeGeneral] || 'other');
    if (it.type === 'other' && /article/i.test(types.resourceType || '')) it.type = 'journalArticle';
    if (it.type === 'other' && /thesis|dissertation/i.test(types.resourceType || '')) it.type = 'thesis';
    it.title = a.titles && a.titles[0] ? trim(a.titles[0].title) : '';
    it.authors = normNames((a.creators || []).map(function (c) {
      if (c.nameType === 'Organizational' || (!c.familyName && !c.givenName)) return { literal: trim(c.name || '') };
      return { family: c.familyName || '', given: c.givenName || '' };
    }));
    it.editors = normNames((a.contributors || []).filter(function (c) { return c.contributorType === 'Editor'; }).map(function (c) { return c.familyName ? { family: c.familyName, given: c.givenName || '' } : { literal: c.name || '' }; }));
    it.year = a.publicationYear || '';
    var d = (a.dates || []).filter(function (x) { return /Issued|Submitted|Created|Available/i.test(x.dateType || ''); })[0];
    if (d) { var pd = parseDateStr(d.date); if (pd) { it.year = it.year || pd.y; if (pd.y === it.year) { it.month = pd.m || ''; it.day = pd.d || ''; } } }
    it.publisher = typeof a.publisher === 'object' && a.publisher ? (a.publisher.name || '') : (a.publisher || '');
    it.doi = cleanDoi(a.doi); it.url = a.url && !/doi\.org/.test(a.url) ? a.url : '';
    it.version = a.version || '';
    it.abstract = a.descriptions && a.descriptions[0] ? stripJats(a.descriptions[0].description) : '';
    it.tags = (a.subjects || []).map(function (x) { return trim(x.subject || ''); }).filter(Boolean).slice(0, 6);
    if (a.container && a.container.title) it.containerTitle = a.container.title;
    if (/^10\.48550\/arxiv\./i.test(it.doi)) { it.type = 'preprint'; it.containerTitle = 'arXiv'; it.number = 'arXiv:' + it.doi.replace(/^10\.48550\/arxiv\./i, ''); it.url = it.url || 'https://arxiv.org/abs/' + it.number.slice(6); }
    if (it.type === 'thesis' && it.publisher) { it.institution = it.publisher; it.publisher = ''; }
    return normalizeItem(it);
  }
  var datacite = {
    lookupDOI: function (doi) {
      doi = cleanDoi(doi);
      return fetchJSON('https://api.datacite.org/dois/' + encodeURIComponent(doi)).then(function (j) { return fromDataCite(j.data && j.data.attributes); });
    }
  };
  var crossref = {
    // Crossref first (journals, books, conferences); DataCite as fallback (arXiv, Zenodo, datasets, software).
    lookupDOI: function (doi) {
      doi = cleanDoi(doi);
      if (!doi) return Promise.reject(new Error('Enter a DOI'));
      return fetchJSON('https://api.crossref.org/works/' + encodeURIComponent(doi)).then(function (j) { return fromCSL(j.message); })
        .catch(function (e) {
          if (!/not found|HTTP 404/.test(e.message || '')) throw e;
          return datacite.lookupDOI(doi).catch(function () { throw new Error('DOI not found in Crossref or DataCite'); });
        });
    },
    search: function (query, rows) {
      return fetchJSON('https://api.crossref.org/works?rows=' + (rows || 10) + '&query.bibliographic=' + encodeURIComponent(query))
        .then(function (j) { return (j.message.items || []).map(fromCSL); });
    },
    arxiv: function (id) { // arXiv IDs are registered as DataCite DOIs (10.48550/arXiv.<id>)
      id = trim(id).replace(/^arxiv:\s*/i, '').replace(/v\d+$/i, '');
      return datacite.lookupDOI('10.48550/arXiv.' + id).then(function (it) { it.type = 'preprint'; it.containerTitle = 'arXiv'; it.number = 'arXiv:' + id; return it; });
    }
  };
  var integrity = {
    // Resolves to { status: 'retracted'|'clean'|'unknown', note, sources:[...] }
    check: function (item) {
      var doi = cleanDoi(item.doi);
      if (!doi) return Promise.resolve({ status: 'unknown', note: 'No DOI — cannot verify automatically', sources: [] });
      var result = { status: 'unknown', note: '', sources: [] }, retracted = false, notes = [];
      var p1 = fetchJSON('https://api.crossref.org/works/' + encodeURIComponent(doi)).then(function (j) {
        var m = j.message || {}; result.sources.push('Crossref');
        (m['updated-by'] || []).forEach(function (u) { if (/retract|withdraw/i.test((u.type || '') + ' ' + (u.label || ''))) { retracted = true; notes.push('Crossref: ' + (u.label || u.type) + (u.DOI ? ' (notice doi:' + u.DOI + ')' : '')); } });
        var t = stripJats(firstOf(m.title)); if (/^(retracted|withdrawn)\b/i.test(t)) { retracted = true; notes.push('Crossref: title marked as retracted'); }
      }, function () {});
      var p2 = fetchJSON('https://api.openalex.org/works/https://doi.org/' + encodeURIComponent(doi) + '?select=id,is_retracted').then(function (j) {
        result.sources.push('OpenAlex');
        if (j && j.is_retracted) { retracted = true; notes.push('OpenAlex (Retraction Watch data): flagged as retracted'); }
      }, function () {});
      return Promise.all([p1, p2]).then(function () {
        if (!result.sources.length) { result.status = 'unknown'; result.note = 'Could not reach Crossref or OpenAlex (offline?)'; return result; }
        result.status = retracted ? 'retracted' : 'clean';
        result.note = retracted ? notes.join('; ') : 'No retraction found (' + result.sources.join(', ') + ')';
        return result;
      });
    }
  };

  /* ------------------------------------------------------------------ */
  /*  Export: RTF and DOCX (minimal, dependency-free)                    */
  /* ------------------------------------------------------------------ */

  function rtfEscape(s) {
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i), ch = s[i];
      if (ch === '\\' || ch === '{' || ch === '}') out += '\\' + ch;
      else if (c < 128) out += ch;
      else if (c > 0xFFFF) { out += '\\u' + (c - 0x10000 > 32767 ? c - 65536 : c) + '?'; }
      else out += '\\u' + (c > 32767 ? c - 65536 : c) + '?';
    }
    return out;
  }
  function htmlToRtfRuns(html) {
    var out = '', parts = String(html).split(/(<\/?(?:i|b|sup)>)/);
    parts.forEach(function (p) {
      if (p === '<i>') out += '{\\i '; else if (p === '</i>') out += '}';
      else if (p === '<b>') out += '{\\b '; else if (p === '</b>') out += '}';
      else if (p === '<sup>') out += '{\\super '; else if (p === '</sup>') out += '}';
      else out += rtfEscape(stripHtml(p));
    });
    return out;
  }
  function exportRtf(entries, opts) {
    opts = opts || {};
    var hanging = !opts.numeric;
    var rtf = '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}}\\f0\\fs24\n';
    if (opts.heading !== false) rtf += '{\\pard\\sa240\\b ' + rtfEscape(opts.heading || 'References') + '\\par}\n';
    entries.forEach(function (e) {
      rtf += '{\\pard' + (hanging ? '\\li720\\fi-720\\sl480\\slmult1' : '\\sa120\\sl360\\slmult1') + ' ' + htmlToRtfRuns(e.html) + '\\par}\n';
    });
    return rtf + '}';
  }

  // --- tiny ZIP writer (stored, no compression) ---
  var CRC_TABLE = (function () { var t = [], c; for (var n = 0; n < 256; n++) { c = n; for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; } return t; })();
  function crc32(bytes) { var c = 0xFFFFFFFF; for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function utf8(str) { if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(str); var b = unescape(encodeURIComponent(str)), a = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; }
  function makeZip(files) { // files: [{name, data(Uint8Array|string)}]
    var parts = [], central = [], offset = 0;
    function u16(n) { return [n & 255, (n >> 8) & 255]; }
    function u32(n) { return [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255]; }
    var dosTime = 0, dosDate = (1 << 5) | 1 | ((2026 - 1980) << 9);
    files.forEach(function (f) {
      var name = utf8(f.name), data = typeof f.data === 'string' ? utf8(f.data) : f.data, crc = crc32(data);
      var local = [].concat(u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(dosTime), u16(dosDate), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0));
      parts.push(new Uint8Array(local), name, data);
      central.push([].concat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(dosTime), u16(dosDate), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset)), name);
      offset += local.length + name.length + data.length;
    });
    var cdStart = offset, cdLen = 0;
    central.forEach(function (c, i) { if (i % 2 === 0) { parts.push(new Uint8Array(c)); cdLen += c.length; } else { parts.push(c); cdLen += c.length; } });
    parts.push(new Uint8Array([].concat(u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length), u32(cdLen), u32(cdStart), u16(0))));
    var total = 0; parts.forEach(function (p) { total += p.length; });
    var out = new Uint8Array(total), pos = 0; parts.forEach(function (p) { out.set(p, pos); pos += p.length; });
    return out;
  }
  function xmlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function htmlToDocxRuns(html) {
    var out = '', parts = String(html).split(/(<\/?(?:i|b|sup)>)/), it = false, b = false, sup = false;
    parts.forEach(function (p) {
      if (p === '<i>') it = true; else if (p === '</i>') it = false;
      else if (p === '<b>') b = true; else if (p === '</b>') b = false;
      else if (p === '<sup>') sup = true; else if (p === '</sup>') sup = false;
      else if (p) {
        var txt = stripHtml(p); if (!txt) return;
        var rpr = (it ? '<w:i/><w:iCs/>' : '') + (b ? '<w:b/>' : '') + (sup ? '<w:vertAlign w:val="superscript"/>' : '');
        out += '<w:r>' + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + xmlEsc(txt) + '</w:t></w:r>';
      }
    });
    return out;
  }
  function exportDocx(entries, opts) {
    opts = opts || {};
    var hanging = !opts.numeric;
    var body = '';
    if (opts.heading !== false) body += '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>' + xmlEsc(opts.heading || 'References') + '</w:t></w:r></w:p>';
    entries.forEach(function (e) {
      var ppr = hanging ? '<w:pPr><w:spacing w:line="480" w:lineRule="auto" w:after="0"/><w:ind w:left="720" w:hanging="720"/></w:pPr>' : '<w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="120"/><w:ind w:left="567" w:hanging="567"/></w:pPr>';
      body += '<w:p>' + ppr + htmlToDocxRuns(e.html) + '</w:p>';
    });
    var doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>';
    var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="240"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/></w:rPr></w:style></w:styles>';
    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>';
    var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    var core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + xmlEsc(opts.heading || 'References') + '</dc:title><dc:creator>LexiRef</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>';
    var app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>LexiRef</Application></Properties>';
    var zip = makeZip([
      { name: '[Content_Types].xml', data: ct }, { name: '_rels/.rels', data: rels },
      { name: 'word/document.xml', data: doc }, { name: 'word/_rels/document.xml.rels', data: docRels }, { name: 'word/styles.xml', data: styles },
      { name: 'docProps/core.xml', data: core }, { name: 'docProps/app.xml', data: app }
    ]);
    if (typeof Blob !== 'undefined') return new Blob([zip], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return zip;
  }
  function exportHtml(entries, opts) {
    opts = opts || {};
    var hanging = !opts.numeric;
    var p = hanging ? 'margin:0 0 0 0.5in;text-indent:-0.5in;line-height:2' : 'margin:0 0 8pt 0;line-height:1.5';
    return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + esc(opts.heading || 'References') + '</title></head><body style="font-family:\'Times New Roman\',serif;font-size:12pt">' + (opts.heading !== false ? '<h1 style="font-size:14pt">' + esc(opts.heading || 'References') + '</h1>' : '') + entries.map(function (e) { return '<p style="' + p + '">' + e.html + '</p>'; }).join('\n') + '</body></html>';
  }

  /* ------------------------------------------------------------------ */
  /*  Sample library                                                     */
  /* ------------------------------------------------------------------ */

  function sampleLibrary() {
    var cols = [{ id: 'c_smartgrid', name: 'Smart grids & microgrids' }, { id: 'c_hydrogen', name: 'Green hydrogen' }, { id: 'c_ml', name: 'Machine learning' }, { id: 'c_methods', name: 'Research methods' }];
    var raw = [
      { type: 'conferencePaper', title: 'Attention is all you need', authors: 'Vaswani, Ashish; Shazeer, Noam; Parmar, Niki; Uszkoreit, Jakob; Jones, Llion; Gomez, Aidan N.; Kaiser, Łukasz; Polosukhin, Illia', year: 2017, containerTitle: 'Advances in Neural Information Processing Systems 30 (NIPS 2017)', place: 'Long Beach, CA, USA', publisher: 'Curran Associates', pages: '5998-6008', url: 'https://papers.nips.cc/paper/7181-attention-is-all-you-need', tags: ['transformers', 'deep learning', 'core'], collections: ['c_ml'], favorite: true, abstract: 'The dominant sequence transduction models are based on complex recurrent or convolutional neural networks that include an encoder and a decoder. We propose a new simple network architecture, the Transformer, based solely on attention mechanisms.' },
      { type: 'journalArticle', title: 'Hydrogen production by PEM water electrolysis – A review', authors: 'Kumar, S. Shiva; Himabindu, V.', year: 2019, containerTitle: 'Materials Science for Energy Technologies', volume: '2', issue: '3', pages: '442-454', doi: '10.1016/j.mset.2019.03.002', tags: ['electrolysis', 'review', 'to-read'], collections: ['c_hydrogen'], notes: 'Good overview of PEM electrolyser materials; use for the Power-to-X lecture.' },
      { type: 'conferencePaper', title: 'MicroGrids', authors: 'Lasseter, Robert H.', year: 2002, containerTitle: '2002 IEEE Power Engineering Society Winter Meeting', place: 'New York, NY, USA', publisher: 'IEEE', volume: '1', pages: '305-308', doi: '10.1109/PESW.2002.985003', tags: ['microgrids', 'core'], collections: ['c_smartgrid'], favorite: true },
      { type: 'journalArticle', title: 'The market value of variable renewables: The effect of solar wind power variability on their relative price', authors: 'Hirth, Lion', year: 2013, containerTitle: 'Energy Economics', volume: '38', pages: '218-236', doi: '10.1016/j.eneco.2013.02.004', tags: ['renewables', 'economics'], collections: ['c_smartgrid'] },
      { type: 'book', title: 'Power electronics: Converters, applications, and design', authors: 'Mohan, Ned; Undeland, Tore M.; Robbins, William P.', year: 2003, edition: '3', publisher: 'John Wiley & Sons', place: 'Hoboken, NJ', isbn: '978-0-471-22693-2', tags: ['textbook', 'power electronics'], collections: ['c_smartgrid'] },
      { type: 'bookSection', title: 'Choices, values, and frames', authors: 'Kahneman, Daniel; Tversky, Amos', editors: 'Kahneman, Daniel; Tversky, Amos', year: 2000, containerTitle: 'Choices, values, and frames', pages: '1-16', publisher: 'Cambridge University Press', place: 'Cambridge', doi: '10.1017/CBO9780511803475.002', tags: ['decision theory'], collections: ['c_methods'] },
      { type: 'report', title: 'The future of hydrogen: Seizing today\u2019s opportunities', authors: '{International Energy Agency}', year: 2019, month: 6, institution: 'International Energy Agency', place: 'Paris', url: 'https://www.iea.org/reports/the-future-of-hydrogen', tags: ['hydrogen', 'policy'], collections: ['c_hydrogen'] },
      { type: 'thesis', title: 'Systems of logic based on ordinals', authors: 'Turing, Alan M.', year: 1938, genre: 'PhD thesis', institution: 'Princeton University', place: 'Princeton, NJ', tags: ['logic', 'classic'], collections: ['c_methods'] },
      { type: 'preprint', title: 'An image is worth 16x16 words: Transformers for image recognition at scale', authors: 'Dosovitskiy, Alexey; Beyer, Lucas; Kolesnikov, Alexander; Weissenborn, Dirk; Zhai, Xiaohua; Unterthiner, Thomas; Dehghani, Mostafa; Minderer, Matthias; Heigold, Georg; Gelly, Sylvain; Uszkoreit, Jakob; Houlsby, Neil', year: 2020, month: 10, containerTitle: 'arXiv', number: 'arXiv:2010.11929', doi: '10.48550/arXiv.2010.11929', tags: ['vision transformers', 'to-read'], collections: ['c_ml'] },
      { type: 'standard', title: 'IEEE Standard for Interconnection and Interoperability of Distributed Energy Resources with Associated Electric Power Systems Interfaces', authors: '{IEEE}', year: 2018, number: 'IEEE Std 1547-2018', publisher: 'IEEE', place: 'New York, NY', doi: '10.1109/IEEESTD.2018.8332112', tags: ['standard', 'DER'], collections: ['c_smartgrid'] },
      { type: 'dataset', title: 'Data package time series', authors: '{Open Power System Data}', year: 2020, version: '2020-10-06', publisher: 'Open Power System Data', doi: '10.25832/time_series/2020-10-06', tags: ['data', 'load profiles'], collections: ['c_smartgrid'] },
      { type: 'webpage', title: 'Ambient (outdoor) air pollution', authors: '{World Health Organization}', year: 2024, month: 10, day: 24, containerTitle: 'World Health Organization', url: 'https://www.who.int/news-room/fact-sheets/detail/ambient-(outdoor)-air-pollution', accessed: todayISO(), tags: ['air quality'], collections: [] },
      { type: 'journalArticle', title: 'Ileal-lymphoid-nodular hyperplasia, non-specific colitis, and pervasive developmental disorder in children', authors: 'Wakefield, A. J.; Murch, S. H.; Anthony, A.; Linnell, J.; Casson, D. M.; Malik, M.; Berelowitz, M.; Dhillon, A. P.; Thomson, M. A.; Harvey, P.; Valentine, A.; Davies, S. E.; Walker-Smith, J. A.', year: 1998, containerTitle: 'The Lancet', volume: '351', issue: '9103', pages: '637-641', doi: '10.1016/S0140-6736(97)11096-0', retracted: true, retractionNote: 'Retracted by The Lancet in 2010 — example of a flagged source.', integrity: { status: 'retracted', checked: new Date().toISOString() }, tags: ['example', 'retracted'], collections: ['c_methods'] }
    ];
    var items = raw.map(function (r, i) { var it = normalizeItem(r); it.id = 'r_' + citeKey(it); it.added = it.modified = new Date(Date.now() - (i + 1) * 60000).toISOString(); return it; });
    return { collections: cols, items: items };
  }

  /* ------------------------------------------------------------------ */
  /*  Export API                                                         */
  /* ------------------------------------------------------------------ */

  root.LexiRef = {
    version: '1.0.1',
    TYPES: TYPES, FIELDS: FIELDS, TYPE_FIELDS: TYPE_FIELDS, STYLES: STYLES, MONTHS: MONTHS,
    uid: uid, esc: esc, stripHtml: stripHtml, cleanDoi: cleanDoi, doiUrl: doiUrl, todayISO: todayISO,
    newItem: newItem, normalizeItem: normalizeItem, parseNames: parseNames, namesToText: namesToText, nameDisplay: nameDisplay,
    formatReference: formatReference, formatCitation: formatCitation, sortForBibliography: sortForBibliography, bibliography: bibliography,
    findDuplicates: findDuplicates, citeKey: citeKey,
    bibtex: { parse: bibtexParse, serialize: bibtexSerialize },
    ris: { parse: risParse, serialize: risSerialize },
    csl: { toCSL: toCSL, fromCSL: fromCSL, parse: cslParse, serialize: function (items) { return JSON.stringify(items.map(toCSL), null, 2); } },
    libraryFile: { serialize: librarySerialize, parse: libraryParse },
    importAny: importAny,
    crossref: crossref, datacite: datacite, integrity: integrity, setContactEmail: setContactEmail,
    exportDocx: exportDocx, exportRtf: exportRtf, exportHtml: exportHtml, makeZip: makeZip,
    sampleLibrary: sampleLibrary
  };
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
