# RA Reply

An Outlook add-in that drafts replies to student research inquiries. You open an email, click **Draft reply**, and an agent decides what kind of inquiry it is and drafts a reply from your playbook. You review and send.

The agent sees only the email you have open. It never sends anything itself.

## Install (once)

1. In a browser, sign in to Outlook on the web and go to <https://aka.ms/olksideload>.
2. Choose **My add-ins** → **Custom Addins** → **Add a custom add-in** → **Add from URL…**
3. Paste `https://www.bingshengyao.com/ra-reply-assistant/manifest.xml` and confirm.

The add-in is now in your mailbox, so it shows up in Outlook for Mac and Outlook for iOS too (it can take a few minutes to appear).

## Set up (once)

Open any email and click **RA Reply**. On a Mac it's on the toolbar, or under **Apps** or **…**. On the first run it asks for:

- **OpenRouter API key.** Use a separate key with a monthly limit.
- **Model.** Defaults to `moonshotai/kimi-k3`. Any OpenRouter model ID works.
- **Playbook.** Paste your playbook (`private/playbook.md` on your Mac). The agent follows it for every draft.

These are stored in your mailbox's settings for this add-in and sync to your other devices. They are not in this repository.

## Use

| Button | What it does |
| --- | --- |
| **Draft reply** | The agent picks the inquiry type from your playbook and drafts the reply it calls for. Above the draft, it says what it decided. |
| **Template buttons** | One per `## Template: …` section in your playbook. You've made the decision; the agent fills in the template for this student. |
| **Note box + Draft from my note** | Type what you want ("say we're full until fall"). After a draft exists, the same box revises it (**Redraft with my note**). A note also applies when you click Draft reply or a template. |
| **Open reply** (Mac) | Opens a normal Outlook reply with the draft filled in. You click Send. |
| **Copy** | Copies the draft. On iPhone, tap Copy, then Reply, and paste. |

**Postdoc inquiries (CC):** Outlook doesn't let add-ins fill in the CC line of a reply. After **Open reply**, click **RA Reply** in the reply window and then **Add colleagues to CC**. On iPhone, add them yourself.

**Pin the panel (Mac):** click the pin icon at the top of the panel to keep it open as you move between emails.

## Edit the playbook

Change it in the add-in's settings (gear icon). The playbook covers:

- **Rules** and **Style**: what the agent must never say, and how it writes.
- **Website pages**: pages it reads for every draft. It refetches them every 6 hours, or right after you save settings.
- **Current details**: recruiting period, interview dates, travel and so on. Update these each round. Blank values are left out of replies.
- **Inquiry types**: what to do for each kind of email when you click **Draft reply**.
- **Templates**: each `## Template: Name` heading becomes a button named *Name*.

Only email addresses written in the playbook can be CC'd.

## Develop

- `node --test test/` runs the tests.
- Preview the panel in a browser without Outlook: serve this folder (`python3 -m http.server`) and open `taskpane.html?dev`. Add `&mock` for canned answers without OpenRouter, `&mobile` for the phone layout, `&compose` for the reply-window view.
- `src/agent.js`: playbook parsing, prompt, OpenRouter call, answer parsing. `src/host.js`: everything that talks to Outlook. `src/app.js`: the panel.
- After changing `manifest.xml`, remove the add-in and add it again from the URL.
