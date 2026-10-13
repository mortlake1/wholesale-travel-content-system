# Pushing this repo to GitHub

Target: **public** repo named `wholesale-travel-content-system`.
The bundle next to this file (`wholesale-travel-content-system.bundle`) contains the full
history, both commits, and the author identity (The Freedom Project, kumbembamj@gmail.com).

## 1. Create the empty repo on GitHub

github.com → **New repository**
- Name: `wholesale-travel-content-system`
- Visibility: **Public**
- Leave every initialisation box **unticked** — no README, no .gitignore, no licence.
  The repo must be completely empty or the push will be rejected.

## 2. Push from your machine

```bash
# from the folder where you saved the bundle
git clone wholesale-travel-content-system.bundle wholesale-travel-content-system
cd wholesale-travel-content-system
git remote set-url origin https://github.com/<your-username>/wholesale-travel-content-system.git
git push -u origin main
```

That's it — four commands, no token needed in this workspace, nothing secret travelled
through chat.

## Notes

- `git clone` accepts a bundle file directly as a source; the clone is a normal repo
  afterwards, and the bundle can be deleted once the push succeeds.
- Commits arrive authored as The Freedom Project (kumbembamj@gmail.com) — GitHub will link
  them to the account that owns that email.
- The workspace copy of this repo stays as the working copy. For future changes, either push
  from your machine as above after re-bundling, or add the remote there once and push directly.
- Public means public: the compliance notes and draft copy will be visible to anyone, including
  competitors. Nothing sensitive is in it (all figures are placeholders), but re-read the README
  once before the push if that ever changes.
