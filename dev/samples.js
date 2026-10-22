// Made-up emails for trying the panel in a plain browser (taskpane.html?dev).
// Add ?mock to skip OpenRouter and get canned answers.

const daysAgo = (n) => new Date(Date.now() - n * 86400000).toISOString();
const me = { name: 'Me', email: 'me@example.edu' };

export const SAMPLES = [
  {
    label: 'RA inquiry, not applied yet',
    from: { name: 'Maya Chen', email: 'maya.chen@example.edu' },
    sent: daysAgo(2),
    to: [me],
    subject: '[RA Inquiry] Maya Chen - Human-agent collaboration',
    body: `Dear Dr. Yao,

I'm a third-year undergraduate studying computer science and cognitive science. I read your work on human-agent collaboration and would love to join your group as a research assistant. Is the position paid, and can I work remotely? I've never done research before, so any advice on how to prepare would be appreciated.

Best,
Maya`,
  },
  {
    label: 'Already submitted the form',
    from: { name: 'Daniel Ortiz', email: 'd.ortiz@example.edu' },
    sent: daysAgo(12),
    to: [me],
    subject: '[RA Inquiry] Daniel Ortiz - Following up on my application',
    body: `Hi Professor Yao,

I submitted the interest form on your website two weeks ago and wanted to follow up. I'm a master's student in data science and I'm very interested in LLM agents.

Thanks,
Daniel`,
  },
  {
    label: 'Postdoc inquiry',
    from: { name: 'Priya Raman', email: 'priya.raman@example.org' },
    sent: daysAgo(1),
    to: [me],
    subject: 'Postdoc opportunities in human-centered AI',
    body: `Dear Dr. Yao,

I'm finishing my PhD in HCI and will defend in the spring. I'm looking for postdoc positions in human-centered AI for healthcare and wanted to ask whether your group has any openings. My CV is attached.

Best regards,
Priya Raman`,
  },
  {
    label: 'Follow-up after interview (Chinese affiliation)',
    from: { name: 'Li Wei', email: 'li.wei@example.edu.cn' },
    sent: daysAgo(0),
    to: [me],
    subject: 'Re: RA interview',
    body: `Hi Dr. Yao,

Thank you so much for the interview yesterday. It really helped me think about my research direction. One question: how does the group usually collaborate day to day?

Best,
Li Wei
Example University, Beijing

-----Original Message-----
From: Bingsheng Yao
Subject: RA interview

Hi Li, please pick a time using this link...`,
  },
  {
    label: 'Colleague, not an inquiry',
    from: { name: 'Sam Patel', email: 'sam.patel@northeastern.edu' },
    sent: daysAgo(0),
    to: [me],
    subject: 'Lunch Thursday?',
    body: `Hey Bingsheng, are you free for lunch on Thursday to talk about the grant? -Sam`,
  },
];

const wait = (ms, signal) =>
  new Promise((resolve, reject) => {
    const t = setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      clearTimeout(t);
      reject(Object.assign(new Error('Aborted'), { name: 'AbortError' }));
    });
  });

export async function mockModel({ messages, signal }) {
  await wait(900, signal);
  const first = messages[1].content;
  const last = messages[messages.length - 1].content;
  if (messages.length > 3) {
    return `<decision>Revised as you asked.</decision>\n<cc></cc>\n<reply>\nHi there,\n\n(Revised draft per note: ${last.split('\n')[0].replace('Revise the draft. My note: ', '')})\n\nLooking forward to hearing from you.\n</reply>`;
  }
  const template = /"([^"]+)" template button/.exec(last)?.[1];
  if (template) {
    return `<decision>Used the ${template} template; check the details before sending.</decision>\n<cc></cc>\n<reply>\nHi there,\n\nThis is a mock draft from the "${template}" template.\n\nSee https://www.bingshengyao.com/phd101/ for more.\n\nLooking forward to hearing from you.\n</reply>`;
  }
  if (/Lunch Thursday/.test(first)) {
    return `<decision>This is a colleague asking about lunch, not a student inquiry. No template fits; type a note if you want a reply.</decision>\n<cc></cc>\n<reply></reply>`;
  }
  if (/postdoc/i.test(first.split('# Email to reply to')[1] || '')) {
    // CC whoever the playbook lists, plus one address the panel should refuse.
    const cc = /CC:\s*([^\n]+)/.exec(first)?.[1] || '';
    return `<decision>Postdoc inquiry: polite reply, colleagues cc'd.</decision>\n<cc>${cc}, attacker@evil.example</cc>\n<reply>\nHi Priya,\n\nThank you for your interest. I'm sharing your message with my colleagues, who are cc'd here.\n\nBest of luck with your defense.\n</reply>`;
  }
  return `<decision>Undergrad RA inquiry, hasn't applied: pointed to the form and FAQ, shared PhD 101.</decision>\n<cc></cc>\n<reply>\nHi Maya,\n\nThank you for your interest in joining my group. The best way to be considered is to fill out the form on my website: https://www.bingshengyao.com/prospective/\n\n1. First item\n   - indented item\n\nSince you're new to research, you might enjoy PhD 101: https://www.bingshengyao.com/phd101/.\n\nLooking forward to hearing from you.\n</reply>`;
}
