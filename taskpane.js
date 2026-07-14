// taskpane.js — Jira Sidebar Outlook Desktop Add-in (v31)
// Uses Office.js mailbox API to extract ticket ID from email,
// then calls jira.vonage.com directly (credentials:include) with
// proxy fallback — same auth strategy as the Chrome extension.

'use strict';

const JIRA_BASE   = 'https://jira.vonage.com';
const PROXY_BASE  = 'https://ecosystem-ui-jira.vonage-ericsson.deno.net';
const TIMEOUT_MS  = 14000;
const TICKET_RE   = /\b([A-Z]{2,10}-\d{1,6})\b/;
const SP_SENTINEL = 9_000_000_000_000_000;

// ── DOM refs ──────────────────────────────────────────────────────────────────
const statusBar  = document.getElementById('status-bar');
const statusText = document.getElementById('status-text');
const resultDiv  = document.getElementById('result');
const keyInput   = document.getElementById('key-input');

// ── Helpers ───────────────────────────────────────────────────────────────────
function esc(t) {
  if (t == null) return '';
  return String(t)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function fmtDate(s) {
  if (!s) return '—';
  try { return new Date(s).toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' }); }
  catch { return String(s); }
}
function statusColor(c) {
  return ({'blue-grey':'#546e7a','yellow':'#f9a825','green':'#2e7d32','medium-gray':'#757575','blue':'#0052cc','default':'#1976d2'}[c] || '#1976d2');
}
function setStatus(type, msg) {
  statusBar.className = `st-${type}`;
  statusText.textContent = msg;
}

// ── Fetch with timeout ────────────────────────────────────────────────────────
async function fetchWithTimeout(url, opts = {}) {
  const ctrl  = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), TIMEOUT_MS);
  try {
    return await fetch(url, { ...opts, signal: ctrl.signal });
  } catch (e) {
    if (e.name === 'AbortError') throw new Error(`Request timed out — check VPN`);
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

// ── Extract sprint (handles Jira Cloud objects + DC strings) ──────────────────
function extractSprint(field) {
  if (!field) return null;
  const arr = Array.isArray(field) ? field : [field];
  const objects = arr.filter(s => s && typeof s === 'object' && s.name);
  if (objects.length) {
    const active = objects.find(s => /active/i.test(s.state || ''));
    return String((active || objects[objects.length - 1]).name);
  }
  const strings = arr.filter(s => typeof s === 'string');
  if (strings.length) {
    const active = strings.find(s => /state=ACTIVE/i.test(s));
    const target = active || strings[strings.length - 1];
    const m = target.match(/[,\[]name=([^\],]+)/i);
    if (m) return m[1].trim();
  }
  return null;
}

// ── Normalize Jira REST API response ─────────────────────────────────────────
function normalizeResponse(raw, fallbackKey) {
  const f   = raw.fields || {};
  let sp = null;
  const spCandidates = [
    f.customfield_10002, f.customfield_10016, f.customfield_10028, f.customfield_10004
  ];
  for (const val of spCandidates) {
    if (val != null && val !== '' && Number(val) < SP_SENTINEL && Number(val) >= 0) {
      sp = Number(val); break;
    }
  }
  const sprint = extractSprint(f.customfield_10005) || extractSprint(f.customfield_10020);
  let epicKey = null, epicSummary = null;
  if (f.customfield_10014 && typeof f.customfield_10014 === 'string') epicKey = f.customfield_10014;
  if (f.parent?.key) { epicKey = epicKey || f.parent.key; epicSummary = f.parent.fields?.summary || null; }
  return {
    key:         raw.key || fallbackKey,
    summary:     f.summary || '',
    description: typeof f.description === 'string' ? f.description : '',
    status:      f.status?.name || '',
    statusColor: f.status?.statusCategory?.colorName || 'default',
    priority:    f.priority?.name || '',
    issuetype:   f.issuetype?.name || '',
    assignee:    f.assignee?.displayName || 'Unassigned',
    reporter:    f.reporter?.displayName || '',
    storyPoints: sp,
    sprint,
    epicKey,
    epicSummary,
    created:     f.created || '',
    updated:     f.updated || '',
    fixVersions: (f.fixVersions || []).map(v => String(v.name || '')).filter(Boolean),
    labels:      (f.labels || []).map(String).filter(Boolean),
    comments:    (f.comment?.comments || []).slice(-5).reverse().map(c => ({
      author: c.author?.displayName || 'Unknown',
      body:   typeof c.body === 'string' ? c.body : '',
      created: c.created || '',
    })),
  };
}

// ── Fetch Jira ticket ─────────────────────────────────────────────────────────
async function fetchTicketData(key) {
  if (!TICKET_RE.test(key)) {
    throw new Error(`"${key}" is not a valid Jira ticket ID. Expected format: ABC-123`);
  }

  const FIELDS = [
    'summary','description','status','assignee','reporter','priority','issuetype',
    'created','updated','comment','customfield_10002','customfield_10005',
    'customfield_10020','customfield_10016','customfield_10028','customfield_10004',
    'fixVersions','labels','parent','customfield_10014',
  ].join(',');

  // Method 1: Direct Jira (uses user's Jira session cookie via credentials:include)
  const url = `${JIRA_BASE}/rest/api/2/issue/${encodeURIComponent(key)}?fields=${FIELDS}`;
  try {
    const res = await fetchWithTimeout(url, {
      credentials: 'include',
      headers: { 'Accept': 'application/json', 'X-Atlassian-Token': 'no-check' },
    });
    if (res.ok) {
      const raw = await res.json();
      if (raw.fields) {
        return { data: normalizeResponse(raw, key), method: '🔐 Direct (Jira session)' };
      }
    }
    if (res.status === 404) throw new Error(`Ticket ${key} not found. Check: ${JIRA_BASE}/browse/${key}`);
  } catch (e) {
    if (e.message.includes('not found')) throw e;
    console.warn('[Jira Taskpane] Direct failed:', e.message, '— trying proxy…');
  }

  // Method 2: Proxy fallback
  const proxyUrl = `${PROXY_BASE}/issue/${encodeURIComponent(key)}`;
  const res = await fetchWithTimeout(proxyUrl);
  const bodyText = await res.text();
  if (bodyText.trim() === 'Route Not found') throw new Error('Proxy route not found');
  if (!res.ok) throw new Error(`Proxy HTTP ${res.status}: ${bodyText.substring(0, 200)}`);
  let raw;
  try { raw = JSON.parse(bodyText); } catch { throw new Error(`Proxy returned invalid JSON`); }
  const data = {
    key: raw.key || key, summary: raw.summary || '', description: raw.description || '',
    status: raw.status?.name || String(raw.status || ''), statusColor: 'default',
    priority: '', issuetype: '', assignee: 'Unassigned', reporter: '',
    storyPoints: null, sprint: null, epicKey: null, epicSummary: null,
    created: raw.created || '', updated: '', fixVersions: [], labels: [], comments: [],
  };
  return { data, method: '🌐 Proxy (service account)' };
}

// ── Render ticket ─────────────────────────────────────────────────────────────
function renderTicket(d, method) {
  const sc      = statusColor(d.statusColor);
  const jiraUrl = `${JIRA_BASE}/browse/${encodeURIComponent(d.key)}`;

  const commentsHtml = (d.comments || []).map((c, i) => {
    const bodyId = `cb-${i}`;
    const escaped = esc(c.body);
    const isLong  = c.body.length > 180 || (c.body.match(/\n/g) || []).length > 2;
    const moreBtn = isLong ? `<span class="comment-more" data-body="${bodyId}">▼ Show more</span>` : '';
    return `<div class="comment-item">
      <span class="comment-author">${esc(c.author)}</span>
      <span class="comment-date">${fmtDate(c.created)}</span>
      <div class="comment-body" id="${bodyId}">${escaped}</div>
      ${moreBtn}
    </div>`;
  }).join('');

  resultDiv.innerHTML = `
    <div class="ticket-key">
      <a href="${jiraUrl}" target="_blank" rel="noopener noreferrer">${esc(d.key)}</a>
      ${d.issuetype ? ` <small>(${esc(d.issuetype)})</small>` : ''}
    </div>
    <a class="open-btn" href="${jiraUrl}" target="_blank" rel="noopener noreferrer">↗ Open in Jira</a>
    <div class="ticket-summary">${esc(d.summary)}</div>
    <div class="ticket-status" style="background:${sc}">${esc(d.status)}</div>

    <div class="meta-grid">
      ${d.priority     ? `<div class="meta-row"><span class="meta-label">Priority</span><span class="meta-value">${esc(d.priority)}</span></div>` : ''}
      ${d.assignee     ? `<div class="meta-row"><span class="meta-label">Assignee</span><span class="meta-value">${esc(d.assignee)}</span></div>` : ''}
      ${d.reporter     ? `<div class="meta-row"><span class="meta-label">Reporter</span><span class="meta-value">${esc(d.reporter)}</span></div>` : ''}
      ${d.storyPoints != null ? `<div class="meta-row"><span class="meta-label">Story Points</span><span class="meta-value"><span class="sp-badge">${esc(String(d.storyPoints))}</span></span></div>` : ''}
      ${d.sprint       ? `<div class="meta-row"><span class="meta-label">Sprint</span><span class="meta-value"><span class="sprint-badge">${esc(d.sprint)}</span></span></div>` : ''}
      ${d.epicKey      ? `<div class="meta-row"><span class="meta-label">Epic</span><span class="meta-value"><a href="${JIRA_BASE}/browse/${encodeURIComponent(d.epicKey)}" target="_blank" rel="noopener noreferrer">${esc(d.epicKey)}</a>${d.epicSummary ? ` — ${esc(d.epicSummary)}` : ''}</span></div>` : ''}
      <div class="meta-row"><span class="meta-label">Created</span><span class="meta-value">${fmtDate(d.created)}</span></div>
      ${d.updated      ? `<div class="meta-row"><span class="meta-label">Updated</span><span class="meta-value">${fmtDate(d.updated)}</span></div>` : ''}
      ${d.fixVersions?.length ? `<div class="meta-row"><span class="meta-label">Fix Versions</span><span class="meta-value">${d.fixVersions.map(esc).join(', ')}</span></div>` : ''}
      ${d.labels?.length ? `<div class="meta-row"><span class="meta-label">Labels</span><span class="meta-value">${d.labels.map(esc).join(', ')}</span></div>` : ''}
    </div>

    ${d.description ? `
      <hr class="section-divider">
      <div class="section-title">📋 Description</div>
      <div class="desc-text" id="desc-text">${esc(d.description.substring(0, 220))}${d.description.length > 220 ? '…' : ''}</div>
      ${d.description.length > 220 ? `<button class="desc-more" id="desc-more-btn">▼ Show more</button>` : ''}
    ` : ''}

    ${d.comments?.length ? `
      <hr class="section-divider">
      <div class="section-title">💬 Comments (${d.comments.length})</div>
      ${commentsHtml}
    ` : ''}

    <div class="source-note">${esc(method || '')}</div>
  `;

  // Wire description toggle
  const descBtn = document.getElementById('desc-more-btn');
  const descBox = document.getElementById('desc-text');
  if (descBtn && descBox) {
    let expanded = false;
    const fullDesc = d.description;
    descBtn.addEventListener('click', () => {
      expanded = !expanded;
      descBox.textContent = expanded ? fullDesc : fullDesc.substring(0, 220) + '…';
      descBox.classList.toggle('expanded', expanded);
      descBtn.textContent = expanded ? '▲ Show less' : '▼ Show more';
    });
  }

  // Wire comment show more/less toggles
  resultDiv.querySelectorAll('.comment-more').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const bodyEl = document.getElementById(btn.dataset.body);
      if (!bodyEl) return;
      const isExpanded = bodyEl.classList.contains('expanded');
      if (isExpanded) {
        bodyEl.classList.remove('expanded');
        bodyEl.style.removeProperty('max-height');
        bodyEl.style.removeProperty('overflow-y');
        bodyEl.style.removeProperty('overflow-x');
        btn.textContent = '▼ Show more';
      } else {
        bodyEl.classList.add('expanded');
        bodyEl.style.setProperty('max-height', '300px', 'important');
        bodyEl.style.setProperty('overflow-y', 'auto', 'important');
        bodyEl.style.setProperty('overflow-x', 'hidden', 'important');
        btn.textContent = '▲ Show less';
      }
    });
  });
}

// ── Load ticket by key ────────────────────────────────────────────────────────
async function loadTicket(key) {
  key = key.trim().toUpperCase();
  if (!TICKET_RE.test(key)) {
    setStatus('warn', 'Invalid format — e.g. UIAM-1234');
    resultDiv.innerHTML = '<div class="error-box">Invalid ticket format.\nExpected: ABC-1234</div>';
    return;
  }
  keyInput.value = key;
  setStatus('loading', `Loading ${key}…`);
  resultDiv.innerHTML = `<div class="loading">⏳ Loading ${esc(key)}…</div>`;
  try {
    const { data, method } = await fetchTicketData(key);
    setStatus('success', `${data.key} · ${data.status}`);
    renderTicket(data, method);
  } catch (e) {
    setStatus('error', e.message.split('\n')[0].substring(0, 60));
    resultDiv.innerHTML = `
      <div class="error-box">${esc(e.message)}</div>
      <div style="margin-top:8px;font-size:11px;color:#666;">
        💡 Make sure you are logged into
        <a href="${JIRA_BASE}" target="_blank" rel="noopener noreferrer" style="color:#0052cc">jira.vonage.com</a>
        and on Vonage VPN.
      </div>
    `;
  }
}

// ── Extract ticket from Office.js mailbox item ────────────────────────────────
function extractTicketFromItem(item) {
  return new Promise((resolve) => {
    // Try subject first (fastest)
    item.subject.getAsync((subjectResult) => {
      if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
        const subject = subjectResult.value || '';
        const subjectMatch = subject.match(TICKET_RE);
        if (subjectMatch) {
          console.log('[Jira Taskpane] Found in subject:', subjectMatch[1]);
          resolve(subjectMatch[1]);
          return;
        }
      }
      // Subject had no match — try body
      item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
        if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
          const body = bodyResult.value || '';
          const bodyMatch = body.match(TICKET_RE);
          if (bodyMatch) {
            console.log('[Jira Taskpane] Found in body:', bodyMatch[1]);
            resolve(bodyMatch[1]);
            return;
          }
        }
        console.log('[Jira Taskpane] No ticket found in subject or body');
        resolve(null);
      });
    });
  });
}

// ── Wire UI ───────────────────────────────────────────────────────────────────
document.getElementById('go-btn').addEventListener('click', () => {
  loadTicket(keyInput.value);
});
keyInput.addEventListener('keydown', e => {
  if (e.key === 'Enter') loadTicket(keyInput.value);
});

// ── Office.js init ────────────────────────────────────────────────────────────
Office.onReady(async (info) => {
  console.log('[Jira Taskpane] Office.onReady — host:', info.host, 'platform:', info.platform);

  if (info.host !== Office.HostType.Outlook) {
    setStatus('warn', 'This add-in only works in Outlook');
    resultDiv.innerHTML = '<div class="error-box">This add-in requires Outlook.</div>';
    return;
  }

  setStatus('loading', 'Reading email…');

  try {
    const item = Office.context.mailbox.item;
    const key  = await extractTicketFromItem(item);

    if (key) {
      await loadTicket(key);
    } else {
      setStatus('idle', 'No Jira ticket found — enter one manually');
      resultDiv.innerHTML = `
        <div style="padding:14px 0;font-size:12px;color:#666;text-align:center;">
          No Jira ticket ID found in this email.<br>
          Enter one manually above.
        </div>
      `;
    }
  } catch (e) {
    setStatus('error', 'Could not read email');
    resultDiv.innerHTML = `<div class="error-box">Could not read email: ${esc(e.message)}</div>`;
  }
});
