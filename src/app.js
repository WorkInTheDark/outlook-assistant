import {
  DEFAULT_MODEL,
  allowedCc,
  buildMessages,
  callModel,
  extractSiteText,
  formatAnswer,
  parseAnswer,
  parsePlaybook,
  revisionMessage,
  toHtml,
} from './agent.js';
import { createHost } from './host.js';

const $ = (id) => document.getElementById(id);
const SITE_CACHE_KEY = 'ra-reply-site-cache';
const SITE_TTL_MS = 6 * 3600 * 1000;
const DRAFTS_KEY = 'ra-reply-drafts';
const LAST_DRAFT_TTL_MS = 30 * 60 * 1000;
const TIMEOUT_MS = 150 * 1000;
const SETTINGS_LIMIT = 30000; // Outlook keeps add-in settings under 32 KB.

let host;
let settings = {};
let playbookInfo = parsePlaybook('');
let email = null; // the open email
let messages = null; // the conversation with the model about this email
let current = null; // the latest { decision, cc, reply }
let abort = null;
let cancelledByUser = false;
let runId = 0;

init().catch((e) => showError(`Couldn't start: ${e.message}`));

async function init() {
  host = await createHost();
  settings = host.loadSettings();
  refreshPlaybook();
  bindEvents();
  if (host.samples) setupDevBar();
  if (host.isMobile) document.body.classList.add('mobile');
  host.onItemChanged(() => loadItem());
  if (!settings.apiKey || !settings.playbook) showSettings(true);
  else await loadItem();
}

function bindEvents() {
  $('settingsBtn').addEventListener('click', () => showSettings(false));
  $('draftBtn').addEventListener('click', () => draft({ kind: 'auto', note: $('note').value }));
  $('noteBtn').addEventListener('click', draftFromNote);
  $('note').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) draftFromNote();
  });
  $('cancelBtn').addEventListener('click', () => {
    cancelledByUser = true;
    abort?.abort();
  });
  $('openReplyBtn').addEventListener('click', openReply);
  $('copyBtn').addEventListener('click', copyDraft);
  $('draft').addEventListener('input', () => {
    autoSize($('draft'));
    if (current) saveDraft({ ...current, reply: $('draft').value });
  });
  $('saveBtn').addEventListener('click', saveSettings);
  $('cancelSettingsBtn').addEventListener('click', () => loadItem());
  $('playbook').addEventListener('input', updateSettingsSize);
  $('addCcBtn').addEventListener('click', addCc);
  $('insertBtn').addEventListener('click', insertDraft);
}

function show(viewId) {
  for (const id of ['mainView', 'composeView', 'settingsView']) $(id).hidden = id !== viewId;
  $('settingsBtn').hidden = viewId === 'settingsView';
  if (viewId === 'settingsView') $('emailMeta').replaceChildren(textEl('strong', 'Settings'));
  if (viewId === 'composeView') $('emailMeta').replaceChildren(textEl('strong', 'Reply helper'));
}

function refreshPlaybook() {
  playbookInfo = parsePlaybook(settings.playbook);
  const box = $('templateBtns');
  box.replaceChildren(
    ...playbookInfo.templates.map((name) => {
      const b = document.createElement('button');
      b.className = 'btn chip';
      b.textContent = name;
      b.addEventListener('click', () => draft({ kind: 'template', name, note: $('note').value }));
      return b;
    }),
  );
  $('templatesLabel').hidden = playbookInfo.templates.length === 0;
}

// ---- The open email -------------------------------------------------------

async function loadItem() {
  stopDrafting();
  resetResult();
  hideError();
  const mode = host.mode();
  if (mode === 'compose') return showCompose();
  show('mainView');
  email = null;
  if (mode === 'none') {
    $('emailMeta').textContent = 'Select an email to reply to.';
    return;
  }
  $('emailMeta').textContent = 'Reading the email…';
  try {
    email = await host.getEmail();
    const who = email.from?.name || email.from?.email || 'Unknown sender';
    $('emailMeta').replaceChildren(textEl('strong', who), textEl('span', email.subject || '(no subject)'));
  } catch (e) {
    $('emailMeta').textContent = '';
    showError(`Couldn't read this email: ${e.message}`);
  }
}

function textEl(tag, text) {
  const el = document.createElement(tag);
  el.textContent = text;
  return el;
}

// ---- Drafting -------------------------------------------------------------

function draft(task) {
  return run(async () => {
    const { pages, failed } = await loadPages(playbookInfo.websiteUrls);
    return { messages: buildMessages({ playbook: settings.playbook, pages, email, task }), failed };
  });
}

function draftFromNote() {
  const note = $('note').value.trim();
  if (!note) {
    $('note').focus();
    showError(current ? 'Type what to change, then click Redraft.' : 'Type what the reply should say first.');
    return;
  }
  if (!current) return draft({ kind: 'note', note });
  const edited = { ...current, reply: $('draft').value };
  const history = messages.slice(0, -1).concat({ role: 'assistant', content: formatAnswer(edited) }, revisionMessage(note));
  return run(async () => ({ messages: history, failed: [] }));
}

async function run(prepare) {
  if (abort) return;
  if (!email) return showError('Open an email first.');
  if (!settings.apiKey || !settings.playbook) return showSettings(true);
  hideError();
  setBusy(true);
  cancelledByUser = false;
  const myRun = runId;
  abort = new AbortController();
  const timer = setTimeout(() => abort?.abort(), TIMEOUT_MS);
  try {
    const { messages: msgs, failed } = await prepare();
    const raw = await (host.mockModel || callModel)({
      apiKey: settings.apiKey,
      model: settings.model || DEFAULT_MODEL,
      messages: msgs,
      signal: abort.signal,
    });
    if (myRun !== runId) return;
    const answer = parseAnswer(raw);
    answer.cc = allowedCc(answer.cc, playbookInfo.emails);
    messages = msgs.concat({ role: 'assistant', content: formatAnswer(answer) });
    $('note').value = '';
    showResult(answer, failed);
  } catch (e) {
    if (myRun !== runId) return;
    if (e.name === 'AbortError') showError(cancelledByUser ? 'Cancelled.' : 'The model took too long to answer. Try again.');
    else showError(e.message);
  } finally {
    clearTimeout(timer);
    abort = null;
    setBusy(false);
  }
}

function stopDrafting() {
  runId += 1; // results of an earlier run no longer apply
  if (abort) {
    cancelledByUser = true;
    abort.abort();
  }
}

let busyTimer = null;
function setBusy(on) {
  for (const b of document.querySelectorAll('#mainView .btn')) b.disabled = on;
  $('busy').hidden = !on;
  clearInterval(busyTimer);
  if (on) {
    const start = Date.now();
    $('busyText').textContent = 'Drafting…';
    busyTimer = setInterval(() => {
      $('busyText').textContent = `Drafting… ${Math.round((Date.now() - start) / 1000)}s`;
    }, 1000);
  }
}

async function loadPages(urls) {
  let cache = {};
  try {
    cache = JSON.parse(localStorage.getItem(SITE_CACHE_KEY) || '{}');
  } catch {}
  const now = Date.now();
  const failed = [];
  const pages = await Promise.all(
    urls.map(async (url) => {
      const hit = cache[url];
      if (hit && now - hit.t < SITE_TTL_MS) return { url, text: hit.text };
      try {
        const res = await fetch(url, { cache: 'no-cache' });
        if (!res.ok) throw new Error(String(res.status));
        const text = extractSiteText(await res.text(), url);
        cache[url] = { t: now, text };
        return { url, text };
      } catch {
        failed.push(url);
        return hit ? { url, text: hit.text } : null;
      }
    }),
  );
  try {
    localStorage.setItem(SITE_CACHE_KEY, JSON.stringify(cache));
  } catch {}
  return { pages: pages.filter(Boolean), failed };
}

// ---- Showing and using the draft -------------------------------------------

function resetResult() {
  current = null;
  messages = null;
  $('result').hidden = true;
  $('draft').value = '';
  $('hint').textContent = '';
  $('noteBtn').textContent = 'Draft from my note';
}

function showResult(answer, failed) {
  current = answer;
  $('result').hidden = false;
  $('decision').textContent = answer.decision || 'Draft ready.';
  $('siteWarning').hidden = failed.length === 0;
  $('siteWarning').textContent = failed.length
    ? `Couldn't read ${failed.join(', ')}, so the draft may be missing details from your website.`
    : '';

  const hasReply = answer.reply.trim().length > 0;
  $('draftBox').hidden = !hasReply;
  $('noteBtn').textContent = hasReply ? 'Redraft with my note' : 'Draft from my note';
  if (!hasReply) {
    current = null;
    $('hint').textContent = 'No draft. Use a template, or type a note to draft one anyway.';
    return;
  }

  $('draft').value = answer.reply;
  autoSize($('draft'));
  $('ccRow').hidden = answer.cc.length === 0;
  $('ccList').textContent = answer.cc.join(', ');
  $('ccHint').textContent = host.isMobile
    ? 'Add them to CC when you paste the reply.'
    : 'After Open reply, open RA Reply in the reply window and click Add colleagues to CC.';
  $('hint').textContent = host.isMobile ? 'Tap Copy, then tap Reply in Outlook and paste.' : '';
  saveDraft(answer);
}

function autoSize(el) {
  el.style.height = 'auto';
  el.style.height = `${Math.min(el.scrollHeight + 2, 640)}px`;
}

async function openReply() {
  hideError();
  const text = $('draft').value;
  try {
    saveDraft({ ...current, reply: text });
    await host.openReply(toHtml(text));
    $('hint').textContent = current.cc.length
      ? 'Reply opened. Add your colleagues to CC, review, and click Send.'
      : 'Reply opened. Review it and click Send.';
  } catch (e) {
    showError(`Couldn't open the reply: ${e.message}. Use Copy instead.`);
  }
}

async function copyDraft() {
  const text = $('draft').value;
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const el = $('draft');
    el.focus();
    el.select();
    document.execCommand('copy');
  }
  $('hint').textContent = host.isMobile ? 'Copied. Tap Reply in Outlook and paste.' : 'Copied.';
}

// Drafts are remembered per conversation so the reply window can add the CCs
// or insert the draft.
function saveDraft(answer) {
  try {
    const drafts = JSON.parse(localStorage.getItem(DRAFTS_KEY) || '{}');
    const entry = { decision: answer.decision, cc: answer.cc, reply: answer.reply, t: Date.now() };
    drafts[email?.conversationId || 'none'] = entry;
    drafts._last = entry;
    const keys = Object.keys(drafts).filter((k) => k !== '_last').sort((a, b) => drafts[b].t - drafts[a].t);
    for (const k of keys.slice(20)) delete drafts[k];
    localStorage.setItem(DRAFTS_KEY, JSON.stringify(drafts));
  } catch {}
}

function findDraft(conversationId) {
  try {
    const drafts = JSON.parse(localStorage.getItem(DRAFTS_KEY) || '{}');
    if (conversationId && drafts[conversationId]) return drafts[conversationId];
    const last = drafts._last;
    return last && Date.now() - last.t < LAST_DRAFT_TTL_MS ? last : null;
  } catch {
    return null;
  }
}

// ---- Reply window -----------------------------------------------------------

let composeDraft = null;

function showCompose() {
  show('composeView');
  composeDraft = findDraft(host.conversationId());
  $('composeHint').textContent = '';
  if (!composeDraft) {
    $('composeDecision').textContent =
      'No draft for this conversation yet. Open RA Reply on the email you are answering and draft there first.';
    $('addCcBtn').hidden = true;
    $('insertBtn').hidden = true;
    return;
  }
  $('composeDecision').textContent = composeDraft.decision || 'Latest draft';
  $('addCcBtn').hidden = composeDraft.cc.length === 0;
  $('composeCc').textContent = composeDraft.cc.length ? `CC: ${composeDraft.cc.join(', ')}` : '';
  $('insertBtn').hidden = !composeDraft.reply;
}

async function addCc() {
  try {
    await host.addCc(composeDraft.cc);
    $('composeHint').textContent = 'Added to CC.';
  } catch (e) {
    $('composeHint').textContent = `Couldn't add CC: ${e.message}`;
  }
}

async function insertDraft() {
  try {
    await host.insertDraft(toHtml(composeDraft.reply), composeDraft.reply);
    $('composeHint').textContent = 'Draft inserted at the top of the reply.';
  } catch (e) {
    $('composeHint').textContent = `Couldn't insert the draft: ${e.message}`;
  }
}

// ---- Settings ---------------------------------------------------------------

function showSettings(firstRun) {
  stopDrafting();
  $('apiKey').value = settings.apiKey || '';
  $('model').value = settings.model || DEFAULT_MODEL;
  $('playbook').value = settings.playbook || '';
  $('settingsIntro').hidden = !firstRun;
  $('cancelSettingsBtn').hidden = !(settings.apiKey && settings.playbook);
  $('settingsError').hidden = true;
  updateSettingsSize();
  show('settingsView');
}

function updateSettingsSize() {
  const n = $('playbook').value.length;
  $('playbookSize').textContent = `${n.toLocaleString()} characters (limit about ${SETTINGS_LIMIT.toLocaleString()})`;
}

async function saveSettings() {
  const values = {
    apiKey: $('apiKey').value.trim(),
    model: $('model').value.trim() || DEFAULT_MODEL,
    playbook: $('playbook').value,
  };
  const fail = (msg) => {
    $('settingsError').textContent = msg;
    $('settingsError').hidden = false;
  };
  if (!values.apiKey || !values.playbook.trim()) return fail('Add your OpenRouter API key and your playbook.');
  if (JSON.stringify(values).length > SETTINGS_LIMIT) {
    return fail("The playbook is too long for Outlook's settings storage. Shorten it a little.");
  }
  $('saveBtn').disabled = true;
  try {
    await host.saveSettings(values);
    settings = values;
    localStorage.removeItem(SITE_CACHE_KEY);
    refreshPlaybook();
    await loadItem();
  } catch (e) {
    fail(`Couldn't save: ${e.message}`);
  } finally {
    $('saveBtn').disabled = false;
  }
}

// ---- Messages -----------------------------------------------------------------

function showError(msg) {
  $('error').textContent = msg;
  $('error').hidden = false;
}

function hideError() {
  $('error').hidden = true;
}

// ---- Dev preview (plain browser only) -------------------------------------------

function setupDevBar() {
  const select = $('devSample');
  host.samples.forEach((s, i) => select.append(new Option(s.label, String(i))));
  select.addEventListener('change', () => host.selectSample(Number(select.value)));
  $('devBar').hidden = false;
}
