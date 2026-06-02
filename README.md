# Geotab SAML Bulk Editor

A MyGeotab add-in that lists active user accounts in a sortable / filterable
table and lets you switch authentication type (`BasicAuthentication` ↔ `SAML`)
and assign / clear SAML issuer certificates in bulk — with inline editing,
a staging buffer, and one-click commit.

Mirrors the UX of [geotab-fuel-transactions-bulk-editor](https://github.com/tacosbembinos/geotab-fuel-transactions-bulk-editor):
the same throttled `multiCall` plumbing, sticky-header table, pending-edits
pill, and modal contract.

## What it touches

Per the [SAML SSO documentation](https://support.geotab.com/mygeotab/doc/sso-saml)
("Additional Options for Enabling SAML Authentication for Users"), switching
a `User` to SAML requires two fields:

```js
user.userAuthenticationType = 'SAML';
user.issuerCertificate      = { id: '<certId>', isRoot: false };
```

Reverting clears `issuerCertificate`. This add-in does exactly that — it
re-`Get`s each user before `Set` (optimistic concurrency on `version`) and
sends the full entity back so untouched fields aren't accidentally nulled.

## Features

- Active-users table with name, first / last, current auth type, assigned
  certificate, last-access timestamp, and active state.
- Inline cell editing on **Auth type** and **SAML certificate** (Enter
  commits, Esc reverts, Tab walks to the next editable cell).
- Per-row **Edit…** modal (full form), bulk **Set SAML…** / **Revert to
  Basic…** actions on the sticky selection bar.
- Pending-edits pill: stages every change client-side; one click commits
  all of them via chunked `multiCall`.
- Search, filter by auth type / certificate / active-only.
- CSV export of the current table or selected rows.
- Rate-limit aware: per-`(method, entity)` sliding-window throttle,
  exponential backoff on `OverLimitException`, single-flight `api.call`
  gate so two callers can't double-emit.
- Cancellation on tab `blur`: aborts in-flight chunks and rejects stale
  callbacks so post-blur side-effects don't clobber fresher state.

## Installation (MyGeotab)

1. **Administration → System… → System Settings → Add-Ins → New Add-In**.
2. Paste the contents of [`manifest.json`](manifest.json) and click **OK**.
3. Save System Settings.
4. The add-in appears under **Administration → SAML Bulk Editor**.

The manifest pins assets to a specific git tag on jsDelivr, so MyGeotab
loads an immutable bundle from the CDN — no per-tenant deployment needed.

## Releasing a new version

Asset URLs are pinned to an immutable tag (NOT `@main`) to prevent
jsDelivr's edge cache from serving stale post-merge content. To cut a
release:

1. Bump the `@vX.Y.Z` tag in **all three** places:
   - `config.json` → `version` + the `url` / `icon` in `items[]`
   - `samlBulkEditor.html` → `<link rel="stylesheet">` and `<script>` src
2. `git commit && git tag vX.Y.Z && git push --tags`
3. Purge jsDelivr's `@main` bootstrap cache (only needed if the manifest's
   `config.json` shape changed):
   - `https://purge.jsdelivr.net/gh/tacosbembinos/geotab-saml-bulk-editor@main/config.json`

## File layout

```
config.json              MyGeotab add-in manifest body (pinned to @vX.Y.Z)
manifest.json            Bootstrap pointer (points to @main/config.json on jsDelivr)
samlBulkEditor.html      Add-in shell, CSP, asset includes
scripts/main.js          Lifecycle, API plumbing, rate-limit infra, table render
styles/main.css          Scoped under .addin-root (do NOT add unscoped rules)
images/icon.svg          Menu icon
```

## Standalone preview

Open `samlBulkEditor.html` from the filesystem. The bottom of `main.js`
detects a non-MyGeotab host and bootstraps with a stub API so the layout,
sort, search, modal, and inline-edit interactions are testable without a
live database. (Loads will fail because there's no real `api.call`.)
