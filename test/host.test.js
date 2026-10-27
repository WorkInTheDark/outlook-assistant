import assert from 'node:assert/strict';
import { test } from 'node:test';

// A stand-in for Outlook shaped like the real Office.js API: settings on
// Office.context.roamingSettings, the open email on Office.context.mailbox.item.
function fakeOffice({ platform = 'Mac' } = {}) {
  const saved = new Map();
  const ok = (value) => ({ status: 'succeeded', value });
  const replies = [];
  const office = {
    HostType: { Outlook: 'Outlook' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    CoercionType: { Text: 'text', Html: 'html' },
    EventType: { ItemChanged: 'olkItemSelectedChanged' },
    onReady: async () => ({ host: 'Outlook', platform }),
    context: {
      diagnostics: { platform, version: '16.0' },
      requirements: { isSetSupported: () => true },
      roamingSettings: {
        get: (k) => saved.get(k),
        set: (k, v) => saved.set(k, v),
        saveAsync: (cb) => cb(ok()),
      },
      mailbox: {
        addHandlerAsync: () => {},
        item: {
          subject: 'RA inquiry',
          conversationId: 'conv-1',
          from: { displayName: 'Maya Chen', emailAddress: 'Maya@Example.edu' },
          to: [{ displayName: 'Me', emailAddress: 'me@example.edu' }],
          cc: [],
          dateTimeCreated: new Date('2026-10-20T12:00:00Z'),
          body: { getAsync: (type, cb) => cb(ok(`body as ${type}`)) },
          displayReplyFormAsync: (data, cb) => {
            replies.push(data);
            cb(ok());
          },
        },
      },
    },
  };
  return { office, saved, replies };
}

test('in Outlook, settings save to and load from roaming settings', async () => {
  const { office, saved } = fakeOffice();
  globalThis.Office = office;
  const { createHost } = await import('../src/host.js');
  const host = await createHost();
  assert.deepEqual(host.loadSettings(), { apiKey: '', model: '', playbook: '' });
  await host.saveSettings({ apiKey: 'k', model: 'm', playbook: 'p' });
  assert.equal(saved.get('apiKey'), 'k');
  assert.deepEqual(host.loadSettings(), { apiKey: 'k', model: 'm', playbook: 'p' });
});

test('in Outlook, the open email is read and a reply opens with the draft', async () => {
  const { office, replies } = fakeOffice();
  globalThis.Office = office;
  const { createHost } = await import('../src/host.js');
  const host = await createHost();
  assert.equal(host.mode(), 'read');
  assert.equal(host.isMobile, false);
  const email = await host.getEmail();
  assert.equal(email.from.email, 'maya@example.edu');
  assert.equal(email.body, 'body as text');
  assert.equal(email.sent, '2026-10-20T12:00:00.000Z');
  await host.openReply('<div>Hi</div>');
  assert.deepEqual(replies, [{ htmlBody: '<div>Hi</div>' }]);
});

test('on iPhone, the host reports mobile', async () => {
  globalThis.Office = fakeOffice({ platform: 'iOS' }).office;
  const { createHost } = await import('../src/host.js');
  assert.equal((await createHost()).isMobile, true);
});
