ERP Workspace UI — STAGING package

Base: staging commit 62bfc56834600b6e6eaf80ce1b42b1180ea3509d

What it adds:
- Persistent internal workspace tabs (session scoped, max 12)
- Open related links in a new ERP screen using Ctrl+Click, middle click, or right click
- Unified right-click menu for links
- Global search shortcut Ctrl+K
- Back/forward/refresh workspace controls
- Status bar and open-screen count
- System-wide visual polish inherited by pages using base.html
- Expanded Chart of Accounts context menu: view/new screen/ledger/edit/add/delete
- Account view quick ledger action
- Search result workspace hints
- Journal view row context fix and cleaner toolbar

Safety:
- UI/templates only. No database schema or accounting posting logic is changed.
- Apply to a branch based on STAGING, test there, then merge only after approval.
- Do not apply directly to main/Production.
