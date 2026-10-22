// Everything that talks to Outlook lives here. Outside Outlook (a plain
// browser), devHost() stands in with sample emails so the panel can be tried.

const SETTING_KEYS = ['apiKey', 'model', 'playbook'];

function promisify(fn) {
  return new Promise((resolve, reject) => {
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message || 'Outlook could not complete that action.'));
    });
  });
}

function person(p) {
  return p ? { name: p.displayName || '', email: (p.emailAddress || '').toLowerCase() } : null;
}

function officeHost() {
  const mailbox = Office.context.mailbox;
  const platform = Office.context.diagnostics?.platform;
  const isMobile = platform === 'iOS' || platform === 'Android';
  const item = () => mailbox.item;
  const isCompose = () => typeof item()?.subject !== 'string';

  return {
    isMobile,
    mode: () => (!item() ? 'none' : isCompose() ? 'compose' : 'read'),

    async getEmail() {
      const it = item();
      const body = await promisify((cb) => it.body.getAsync(Office.CoercionType.Text, cb));
      return {
        conversationId: it.conversationId || '',
        from: person(it.from || it.sender),
        sent: it.dateTimeCreated ? new Date(it.dateTimeCreated).toISOString() : null,
        to: (it.to || []).map(person),
        cc: (it.cc || []).map(person),
        subject: it.subject || '',
        body,
      };
    },

    conversationId: () => item()?.conversationId || '',

    async openReply(html) {
      const it = item();
      if (Office.context.requirements.isSetSupported('Mailbox', '1.9')) {
        await promisify((cb) => it.displayReplyFormAsync({ htmlBody: html }, cb));
      } else {
        it.displayReplyForm({ htmlBody: html });
      }
    },

    async addCc(emails) {
      await promisify((cb) => item().cc.addAsync(emails, cb));
    },

    async insertDraft(html, text) {
      const body = item().body;
      const type = await promisify((cb) => body.getTypeAsync(cb));
      const isHtml = type === Office.CoercionType.Html;
      await promisify((cb) =>
        body.prependAsync(isHtml ? `${html}<div><br></div>` : `${text}\n\n`, { coercionType: type }, cb),
      );
    },

    loadSettings() {
      const rs = mailbox.roamingSettings;
      return Object.fromEntries(SETTING_KEYS.map((k) => [k, rs.get(k) || '']));
    },

    async saveSettings(values) {
      const rs = mailbox.roamingSettings;
      for (const k of SETTING_KEYS) rs.set(k, values[k] || '');
      await promisify((cb) => rs.saveAsync(cb));
    },

    onItemChanged(handler) {
      mailbox.addHandlerAsync?.(Office.EventType.ItemChanged, handler);
    },
  };
}

function devHost(samples, mockModel) {
  let index = 0;
  const store = 'ra-reply-dev-settings';
  const listeners = [];
  return {
    isMobile: new URLSearchParams(location.search).has('mobile'),
    mode: () => (new URLSearchParams(location.search).has('compose') ? 'compose' : 'read'),
    samples,
    selectSample(i) {
      index = i;
      listeners.forEach((fn) => fn());
    },
    async getEmail() {
      return { ...samples[index], conversationId: `sample-${index}` };
    },
    conversationId: () => `sample-${index}`,
    async openReply(html) {
      const w = window.open('', '_blank');
      if (w) w.document.body.innerHTML = `<p style="color:#666;font:13px sans-serif">Reply preview (dev mode)</p>${html}`;
    },
    async addCc(emails) {
      console.log('[dev] add CC', emails);
    },
    async insertDraft(html) {
      console.log('[dev] insert draft', html);
    },
    loadSettings() {
      try {
        return { apiKey: '', model: '', playbook: '', ...JSON.parse(localStorage.getItem(store) || '{}') };
      } catch {
        return { apiKey: '', model: '', playbook: '' };
      }
    },
    async saveSettings(values) {
      localStorage.setItem(store, JSON.stringify(values));
    },
    onItemChanged(handler) {
      listeners.push(handler);
    },
    mockModel,
  };
}

export async function createHost() {
  if (typeof Office !== 'undefined') {
    const info = await Office.onReady();
    if (info.host === Office.HostType.Outlook && Office.context.mailbox) return officeHost();
  }
  const dev = await import('../dev/samples.js');
  const useMock = new URLSearchParams(location.search).has('mock');
  return devHost(dev.SAMPLES, useMock ? dev.mockModel : null);
}
