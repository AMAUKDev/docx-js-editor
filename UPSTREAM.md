# Upstream lineage, mirrors, and cherry-pick guide

This is a **fork**. This document records where it came from, why the GitHub
"forked from" label changed, where offline mirrors live, and how to pull
upstream fixes (and why it is harder than it looks).

Captured 2026-07-20.

---

## TL;DR

- Our fork: `https://github.com/AMAUKDev/docx-js-editor` (flat single-package layout).
- Our **original parent** `eigenpal/docx-js-editor` was **deleted** — GitHub re-pointed
  the "forked from" label to a surviving sibling, `DoctorSlimm/docx-js-editor`.
- Upstream restructured **flat → monorepo** one day after our divergence point, and
  rewrote history. So almost all upstream fixes now live in a relocated `packages/*`
  layout that does **not** match our flat `src/` tree → cherry-picking is manual porting,
  not clean picks.
- Offline mirrors are kept in `Dropbox\ZakHodgson\Packages` (see below).

---

## Remote configuration

### Original config (before any changes this session)

```
origin    https://github.com/AMAUKDev/docx-js-editor.git   (fetch & push)
upstream  https://github.com/eigenpal/docx-editor          (fetch & push)
```

No `doctorslimm` remote existed. `upstream` points at `eigenpal/docx-editor`
(the monorepo, no `-js-`), NOT the deleted `eigenpal/docx-js-editor`.

> Note: during investigation on 2026-07-20 the `upstream` remote was temporarily
> repointed at a local `C:\` mirror and a `doctorslimm` remote was added, then
> **restored to the original config above**. Remotes live in `.git/config` and are
> not committed, so this never entered git history.

### Why a colleague cloning our fork cannot reach upstream

`upstream` is **local `.git/config`** — it does **not** travel with a clone.
A fresh clone of `origin` has only `origin`. Each person must add `upstream`
themselves. If they previously added it pointing at `eigenpal/docx-js-editor`
(with `-js-`), that now 404s because that repo was deleted.

### Colleague setup — copy/paste

```bash
cd path/to/docx-js-editor
git remote -v                              # see what you currently have

git remote remove upstream 2>/dev/null     # clear any stale/deleted pointer
git remote add upstream https://github.com/eigenpal/docx-editor.git   # active monorepo
git remote add doctorslimm https://github.com/DoctorSlimm/docx-js-editor.git  # flat baseline (optional)

git fetch upstream
git fetch doctorslimm
git remote -v
git log -1 --oneline upstream/main         # should show a recent eigenpal commit
```

> Do **NOT** use `github.com/eigenpal/docx-js-editor` (with `-js-`) — deleted, 404s.
> The live upstream is `eigenpal/docx-editor` (no `-js-`).

---

## Offline mirrors (for safe-keeping)

Location: `I:\Dropbox (Personal)\ZakHodgson\Packages\docx-js-editor\upstream\`

| Artifact | What it is |
| --- | --- |
| `eigenpal-docx-editor.git\` | Full `--mirror` of `eigenpal/docx-editor` (active monorepo upstream, ~59M). |
| `doctorslimm-docx-js-editor.git\` | Full `--mirror` of `DoctorSlimm/docx-js-editor` (flat frozen baseline, ~61M). |
| `snapshots\eigenpal-last-flat-262a1d01.bundle` | Git bundle of the **last flat-layout upstream commit** (`262a1d01`), before the monorepo extraction. Matches our `src/` layout. |

These are bare repos / a bundle. To back them up, just copy the folders/file.
To use a mirror or bundle:

```bash
# fetch from a bare mirror
git remote add upstream-mirror "I:/Dropbox (Personal)/ZakHodgson/Packages/docx-js-editor/upstream/eigenpal-docx-editor.git"
git fetch upstream-mirror

# fetch from the last-flat bundle
git fetch "I:/Dropbox (Personal)/ZakHodgson/Packages/docx-js-editor/upstream/snapshots/eigenpal-last-flat-262a1d01.bundle" upstream-last-flat
```

A local tag `upstream-last-flat -> 262a1d01` was created in this working repo to
produce the bundle. Tags are not pushed unless you push them.

---

## Lineage / key commits

```
aefc0c64  2026-02-08  DoctorSlimm/main HEAD — FROZEN. Our direct ancestor.
                      (our fork = DoctorSlimm's full history + 236 commits on top;
                       DoctorSlimm has 0 commits we lack.)
   |
 ...shared dev...
   |
d8d43207  2026-03-02  Last commit common to our fork AND eigenpal/0.x.
262a1d01  2026-03-03  LAST FLAT-LAYOUT upstream commit (bundled above).
e1ada8db  2026-03-03  "extract monorepo" — upstream goes flat -> packages/*.
   |
   +--> eigenpal/main   monorepo, history-rewritten, active (HEAD bdb3ccc, 2026-07-20)
   +--> our fork        stayed flat, HEAD ~2026-07-13
```

---

## Cherry-pick verdict

| Source | Newer than us? | Commits we lack | Layout | Cherry-pick |
| --- | --- | --- | --- | --- |
| `DoctorSlimm/main` | No — frozen 2026-02-08 | **0** | flat | Nothing to pull. Baseline only. |
| `eigenpal/0.x` and `main` | Yes | ~244 (102 fixes) | `packages/{core,react}/src/` | Manual porting — see below. |

The monorepo extraction landed **one day after** our divergence point, so the clean
(flat, path-matching) window is only ~5 commits / 1 fix. All ~102 valuable fixes are
on the monorepo side and touch relocated paths that fold two upstream prefixes
(`packages/core/src`, `packages/react/src`) into our single `src/`.

### Porting a monorepo-side fix onto our flat fork

```bash
git show <sha> -- packages/core packages/react \
  | sed 's#packages/core/src#src#; s#packages/react/src#src#' \
  | git apply --3way
```

Then hand-reconcile anywhere our internal structure (`src/docx`, `src/prosemirror`,
`src/layout-*`, `src/components`) does not line up with theirs. Do this per-fix, not
wholesale. High-value candidates (map to our DOCX-preservation risks): SDT preservation
(#482, #486), header/footer refs (#481, #416), dense footnotes (#485), OOXML integer
coercion (#422).
