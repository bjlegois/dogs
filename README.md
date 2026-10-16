# Let-Outs

Bubba and Harvey's let-out board: https://bjlegois.github.io/dogs/

Open it on a phone and use **Add to Home Screen** so it opens like an app. The first time, it asks for the family passcode and who you are.

## How it decides let-outs

- **Bubba** (Bryce & McKayla's house): 1 let-out on each day Bryce and McKayla both work. None if either is off.
- **Harvey** (Brynn's house): 1 let-out on any day Brynn works, 2 when she works 12 hours or more. Days Brynn hasn't entered show as "hours not entered".
- Anyone can claim a let-out. Unclaimed ones show in red with who's free that day. Bryce and McKayla count as free for Harvey's 2nd let-out.
- Any dog's count can be changed by hand for a single day (**Change** on the day card).

## Where the data lives

Everything is in a private Google Sheet ("Let-Out Board") with an Apps Script web app behind it. This repo holds only the app. The passcode is stored in the Sheet's script, never here.

Sheet tabs: `People`, `Dogs`, `Hours`, `Claims`, `Counts`, and `Coverage`.

`Coverage` lists every let-out from today through the end of next week, with who claimed it or `UNCLAIMED`, plus days whose hours aren't entered. It is rebuilt on every change and every night at 1 AM, so it can be read on any schedule.

## Backend setup (one time)

1. Create a new Google Sheet and name it "Let-Out Board".
2. Extensions > Apps Script. Replace everything with `backend/Code.gs`, and put the passcode in `PASSCODE` at the top.
3. Pick `setup` in the function menu and click Run. Approve the permissions.
4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Copy the Web app URL.
5. Put that URL in `config.js`.

After editing `Code.gs` later, use Deploy > Manage deployments > Edit > New version so the same URL picks up the change.
