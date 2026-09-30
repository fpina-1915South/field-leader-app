# 1915 Field Leader App

One app for Market Leaders and directors: it tells each field leader where to go and what to coach, then walks them through the 6 Elements visit in the store. Every Sunday it builds a 5-visit week from Saturday's numbers. After each daily upload it suggests a swap if a store gets worse. No IT needed: GitHub Pages hosts the app, and Firebase (free plan) handles sign-in and saving.

## Go live

Firebase is already set up: project `field-leader-1915`, Email/Password sign-in on, Firestore created (US, production mode) with `firestore.rules` published, and the web app's keys are in `config.js`. What's left:

1. **Put it on GitHub.** Create a new repository called `field-leader-app` and upload all the files in this folder. Then go to Settings > Pages, set Source to "Deploy from a branch", pick `main` and `/ (root)`, and click Save. After a minute your link is `https://<your-github-name>.github.io/field-leader-app/`.
2. **Allow the link to sign in.** In Firebase, go to Authentication > Settings > Authorized domains and add `<your-github-name>.github.io`.
3. **Sign in first.** Open the link and click Create account with fpina@1915south.com. Verify your email, and you're in as admin.
4. **Set up markets and leaders.** In Setup, add each Market Leader and director under Logins with their email, role, stores and default days off (any 2 days; Wed and Thu if you don't pick). Then create your Markets: name each one, pick its stores, and name its Market Leader and director. Their store lists follow the market, and a store belongs to one market. Each leader creates their own password with the same email and picks their days off for any week on My week; the schedule builds as soon as they save.
5. **Upload.** In Upload, drop in the sales team roster and the store leader list (email, name, role, stores) once, and again whenever they change. Then every morning, drop in the daily report and the RSA report together. The RSA report needs to go in daily, including Saturday: the app compares each day to the day before (the daily brief) and to last Saturday's copy (this week).

## Weekly 1 on 1s

Admins hold a weekly 1 on 1 with each Market Leader on the **1 on 1s** tab, covering the week that just closed (Sunday to Saturday). It builds itself from that week's last daily report and RSA report: wins and opportunities, who performed and who didn't, how the leader ran their visits, the lever the market needs to pull, and where to focus this week. It ends with up to 3 commitments (from X to Y by a date, and how). The Market Leader sees them on My 1 on 1 and on their daily brief, and the next 1 on 1 opens by reviewing them. Your private notes are stored separately and only admins and the exec team can read them.

## Files

- `index.html`, `app.js`: the app
- `ml.js`: need score, Sunday plan, pivot, consultant week-over-week, and the 6 Elements visit content and scoring (from the director Store Visit app)
- `msgs.js`: the Team messages (daily huddle, daily market recap, weekly kickoff, weekly market update, visit recap)
- `base.js`: report parsing and coaching talk tracks, shared with the Consultant Scorecard
- `config.js`: the Firebase keys for `field-leader-1915`. This web key is meant to be public; `firestore.rules` controls who can read and write. If `apiKey` starts with PASTE, the app runs in demo mode with sample data.
- `firestore.rules`: who can read and write what

Photos taken on a visit are shrunk on the phone and saved in Firestore (20 per visit, each tagged to an element with comments), so the free plan covers them. The free plan holds 1 GB, which is roughly 10,000 visit photos; Frank can delete old photos if it ever gets close. The Store Visit app's Power Automate submit is no longer needed: visits save straight to the app.

## The daily report export

The export needs these columns: `Report Date, Segment, Metric`, plus `MTD TY / MTD LY / MTD Budget`. Keep `WTD TY / WTD LY / WTD Budget` in the export too so the plan can react to the current week. Without them, it runs on month-to-date numbers only.
