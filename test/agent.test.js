import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { test } from 'node:test';
import {
  allowedCc,
  buildMessages,
  callModel,
  extractSiteText,
  formatAnswer,
  parseAnswer,
  parsePlaybook,
  revisionMessage,
  toHtml,
} from '../src/agent.js';

const PLAYBOOK = `# Playbook
## Rules
- Be nice.

## Website pages
- https://example.com/
- https://example.com/prospective/

## Inquiry types
### Postdoc inquiry
### Not a section
- CC: a.one@northeastern.edu, B.Two@northeastern.edu

## Template: Schedule interview
Hi {first name},

## Template:   Reject after interview
Hi {first name},
`;

test('parsePlaybook finds templates, website pages, and CC-able addresses', () => {
  const info = parsePlaybook(PLAYBOOK);
  assert.deepEqual(info.templates, ['Schedule interview', 'Reject after interview']);
  assert.deepEqual(info.websiteUrls, ['https://example.com/', 'https://example.com/prospective/']);
  assert.ok(info.emails.has('a.one@northeastern.edu'));
  assert.ok(info.emails.has('b.two@northeastern.edu'));
});

test('the real playbook parses into three templates and four pages', { skip: !existsSync('private/playbook.md') }, () => {
  const info = parsePlaybook(readFileSync('private/playbook.md', 'utf8'));
  assert.deepEqual(info.templates, ['Schedule interview', 'Accept after interview', 'Reject after interview']);
  assert.equal(info.websiteUrls.length, 4);
  assert.ok(info.emails.size >= 2, 'the postdoc CC addresses are listed');
});

test('parseAnswer reads the tagged format', () => {
  const a = parseAnswer('<decision>Postdoc.</decision>\n<cc>X@Y.edu, x@y.edu</cc>\n<reply>\nHi Priya,\n\nThanks.\n</reply>');
  assert.deepEqual(a, { decision: 'Postdoc.', cc: ['x@y.edu'], reply: 'Hi Priya,\n\nThanks.' });
});

test('parseAnswer copes with thinking blocks, missing closing tags, and template separators', () => {
  const a = parseAnswer('<think>hmm</think><decision>RA</decision><cc></cc><reply>---\nHi Maya,\n\nThanks.\n---');
  assert.equal(a.decision, 'RA');
  assert.equal(a.reply, 'Hi Maya,\n\nThanks.');
});

test('parseAnswer treats untagged text as the reply', () => {
  assert.deepEqual(parseAnswer('Hi Sam,\n\nSure.'), { decision: '', cc: [], reply: 'Hi Sam,\n\nSure.' });
});

test('an empty reply means "do not draft"', () => {
  const a = parseAnswer('<decision>Colleague email.</decision><cc></cc><reply></reply>');
  assert.equal(a.reply, '');
});

test('allowedCc drops addresses that are not in the playbook', () => {
  const { emails } = parsePlaybook(PLAYBOOK);
  assert.deepEqual(allowedCc(['a.one@northeastern.edu', 'attacker@evil.example'], emails), ['a.one@northeastern.edu']);
});

test('formatAnswer round-trips through parseAnswer', () => {
  const answer = { decision: 'd', cc: ['a@b.co'], reply: 'Hi,\n\nBody.' };
  assert.deepEqual(parseAnswer(formatAnswer(answer)), answer);
});

test('toHtml escapes, links URLs without trailing punctuation, and keeps paragraphs and indentation', () => {
  const html = toHtml('Hi <Maya> & co,\n\nSee https://x.com/a?b=1&c=2.\n   - item');
  assert.equal(
    html,
    '<div>Hi &lt;Maya&gt; &amp; co,</div><div><br></div>' +
      '<div>See <a href="https://x.com/a?b=1&amp;c=2">https://x.com/a?b=1&amp;c=2</a>.</div>' +
      '<div>&nbsp;&nbsp;&nbsp;- item</div>',
  );
});

test('extractSiteText drops scripts and navigation and keeps link targets', () => {
  const html = `<html><nav><a href="/cv/">CV</a></nav><script>var x=1</script>
    <article><h1>Prospective Students</h1><p>Fill out <a href="https://forms.gle/abc">this form</a> &amp; read
    the <a href="/norms/">lab norms</a>.</p><p>It&#8217;s remote.</p></article><footer>©</footer></html>`;
  const text = extractSiteText(html, 'https://example.com/prospective/');
  assert.equal(
    text,
    'Prospective Students\nFill out this form (https://forms.gle/abc) & read the lab norms (https://example.com/norms/).\nIt’s remote.',
  );
});

test('buildMessages includes the playbook, website, email age, and task', () => {
  const now = new Date('2026-10-22T12:00:00Z');
  const [system, user] = buildMessages({
    playbook: PLAYBOOK,
    pages: [{ url: 'https://example.com/', text: 'About me' }],
    email: {
      from: { name: 'Maya Chen', email: 'maya@example.edu' },
      sent: '2026-10-12T12:00:00Z',
      to: [{ name: 'Me', email: 'me@x.edu' }],
      cc: [],
      subject: 'RA',
      body: 'Hello',
    },
    task: { kind: 'template', name: 'Schedule interview', note: 'mention spring' },
    now,
  });
  assert.equal(system.role, 'system');
  assert.match(system.content, /<decision>/);
  assert.match(user.content, /# Playbook\n\n# Playbook/);
  assert.match(user.content, /## https:\/\/example.com\/\nAbout me/);
  assert.match(user.content, /From: Maya Chen <maya@example.edu>/);
  assert.match(user.content, /\(10 days ago\)/);
  assert.match(user.content, /"Schedule interview" template button/);
  assert.match(user.content, /My note: mention spring/);
});

test('revisionMessage carries the note', () => {
  assert.match(revisionMessage('  shorter  ').content, /My note: shorter\n/);
});

test('callModel sends the request OpenRouter expects and returns the text', async () => {
  let sent;
  const fetchImpl = async (url, init) => {
    sent = { url, init };
    return { ok: true, status: 200, json: async () => ({ choices: [{ message: { content: 'OK' } }] }) };
  };
  const out = await callModel({ apiKey: 'k', model: 'm/x', messages: [{ role: 'user', content: 'hi' }], fetchImpl });
  assert.equal(out, 'OK');
  assert.equal(sent.url, 'https://openrouter.ai/api/v1/chat/completions');
  assert.equal(sent.init.headers.Authorization, 'Bearer k');
  assert.equal(JSON.parse(sent.init.body).model, 'm/x');
});

test('callModel explains a bad key', async () => {
  const fetchImpl = async () => ({ ok: false, status: 401, json: async () => ({ error: { message: 'No auth' } }) });
  await assert.rejects(callModel({ apiKey: 'k', messages: [], fetchImpl }), /rejected the API key/);
});
