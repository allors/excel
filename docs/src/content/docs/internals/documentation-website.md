---
title: "Documentation website"
description: "The adoption of the Allors documentation standard for Allors Excel, what was reused, how it was validated, and where the website is published."
audience: "Maintainers"
kind: "Decision record"
status: "Current"
sidebar:
  order: 4
---

> **Status: Current.** The website, its Markdown exports and the checks that the documentation
> workflow runs are built and pass for both deployment paths, and the website is published at
> `https://docs.allors.com/excel/`, beneath the Allors documentation home, from the repository's
> Pages source **GitHub Actions**.

This record describes the documentation website of Allors Excel: what was decided, what was
reused from the other Allors projects and what remains to do. The
[writing guide](writing-documentation.md) owns the editorial rules and
[Terminology](../terminology.md) owns the vocabulary. The pages live in `docs/src/content/docs/`;
the [website instructions](../../../../README.md) cover local development, validation and
publication.

## Decision

Allors Excel adopts the Allors documentation standard as first implemented for Allors3 in
[allors3 pull request #284](https://github.com/allors/allors3/pull/284) and adopted since by
Immersive, in [immersive pull request #1](https://github.com/allors/immersive/pull/1), and by
Allors Documents, in [documents pull request #9](https://github.com/allors/documents/pull/9). The
Allors documentation home is intended to present the Allors platforms and a collection of
supporting open-source tools, each product publishing its own documentation beneath its own path.
Allors Excel is one of the supporting tools; its documentation is smaller than a platform's and
proportional to its scope, with the same writing and accuracy standard.

| Area | Decision |
| --- | --- |
| Readers | Human documentation serves users, who write programs and host them in Excel or run them headless, and maintainers. Agent instructions in `AGENTS.md` point to the same content. |
| Site | Astro with Starlight. `docs/` is the website project, with the content in `docs/src/content/docs/`. User and maintainer content have separate navigation groups. |
| Authoring | Markdown. MDX and components are available when a page needs them; none is used yet. |
| Hosting | GitHub Pages, built with GitHub Actions, at the path `/excel/` beneath the shared origin `https://docs.allors.com`, whose root is the Allors documentation home, with `/` also validated. The sidebar's first entry links to the home, and the link checker leaves that link to the home's own build. The origin's own site and its custom domain belong to the documentation home, not to this repository. |
| Version context | Pages and exports identify the documented version as `main · 1.1 (unreleased)`: the version from `version.json`, which Nerdbank.GitVersioning stamps on the packages, and the fact that releases are built from tags, so that `main` is unreleased. |
| Visuals | Mermaid where suitable; the diagrams in [How the virtual workbook works](../virtual-workbook.md) and [Architecture](architecture.md) are the first. No interactive component yet. |
| Language and structure | Oxford English, precise practitioner terminology and Diátaxis, as defined in the writing guide. |
| Checks | The tooling tests, the Astro check, the production build with exports, the link check and browser tests with accessibility checks, for both deployment paths. |

## What was reused

The files were copied from the Allors Documents adoption at commit
`62794a7`, the merge of its pull request #9 on `main`, inspected on 2026-10-19, because that copy
carries the fixes made since the Allors3 original: the link checker names files with forward
slashes on every platform, the `focusable-code-blocks.mjs` plugin gives every code block its
`tabindex` at build time, and the version label is read from `version.json`. The project-specific
changes:

| File | Change for Allors Excel |
| --- | --- |
| `astro.config.mjs` | Sidebar entries; the title and description are read from `site.mjs`. |
| `site.mjs` | Title, description, repository, the version label built from `version.json` and the default base path `/excel/`. |
| `integrations/exports.mjs` | The agent index heading and description are read from `site.mjs`, so the file contains no project text; the audience order is that of Immersive. Immersive read the title only and Allors Documents wrote both into the file; port this form to both. |
| `integrations/source-links.mjs`, `integrations/scrollable-content.mjs`, `integrations/focusable-code-blocks.mjs` | Unchanged. |
| `scripts/check-links.mjs` | The default base path. |
| `src/content.config.ts` | The audience names of Immersive: `Users`, `Maintainers` and `Users and maintainers`. |
| `src/components/PageTitle.astro`, `src/styles/documentation.css`, `public/favicon.ico`, `tsconfig.json`, `playwright.config.ts`, `.gitignore` | Unchanged; the favicon is the one of Allors3 and Immersive. |
| `package.json`, `package-lock.json` | The package name; the pinned dependency versions are identical. |
| `tests/exports.test.mjs`, `tests/links.test.mjs`, `tests/source-links.test.mjs`, `tests/fixtures/anchors.md` | Base paths, repository paths, version strings and the audience names. |
| `tests/browser/documentation.spec.ts` | The page names, the diagram and the table it exercises. |
| `.github/workflows/documentation.yml` | The base path matrix and `version.json` as a trigger path, because the version label depends on it. The `CI` workflow gained `paths-ignore` for `docs/`, the documentation workflow and `AGENTS.md`, as Immersive's has, because its two legs install Office and take a quarter of an hour. |

This third adoption confirms the shared set the Immersive and Allors Documents records describe:
the integrations, the link checker, the title component, the stylesheet, the content schema, the
tooling tests and the workflow, parametrized by `site.mjs`. With the title and description read
from `site.mjs`, `exports.mjs` no longer differs between projects except for the audience order,
which the content schema also holds; moving the audiences into `site.mjs` would make both files
identical. Extracting the set is a decision for the Allors documentation home; until then, keep the
copies identical where the table says so, and port fixes between the repositories.

## Validation

On 2026-10-19 the following ran locally from `docs/` for the base paths `/excel/` and `/`:
`npm test`, `npm run check`, `npm run build`, `npm run check:links` and `npm run test:browser`,
as the [website instructions](../../../../README.md) describe. The tutorial in
[Getting started](../getting-started.md) was run from a clean directory against a clone of the
repository, and the repository's build and tests were run as [Build and test](build-and-test.md#local-run)
records, including the failure of the Excel passes on a machine whose Excel shows a modal dialog.

Automated checks do not establish that an explanation is accurate or clear. The writing guide
requires editorial review of changed pages and of representative exports.

## Remaining work

- Decide with the other Allors repositories whether to extract the reused tooling.
- Release from `main`, then update the README and [Packages and versions](../packages.md) so that
  they no longer say the published packages predate the API.
