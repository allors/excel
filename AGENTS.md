# AGENTS.md

Instructions for agents working in this repository. The technical facts live in the human
documentation; this file says where to find them and how to work here.

## Scope

- Allors Excel is a .NET library for Excel add-ins. A program implements `IProgram` and works
  with a virtual workbook: workbooks, worksheets and cells with zero-based coordinates, whose
  changes are written to Excel in blocks when a worksheet is flushed. `Allors.Excel.Headless`
  keeps the workbook in memory; `Allors.Excel.Interop` runs inside Excel through COM, created by
  an Excel-DNA or VSTO host. [README.md](README.md) introduces the project and the
  [documentation index](docs/src/content/docs/index.md) lists the pages.
- Allors Excel is an open-source supporting tool of Allors. Public documentation, code comments,
  commit messages and GitHub text describe Allors Excel on its own. Do not name or link private,
  unpublished Allors projects.

## Development notes

- The solution is `src/Excel.sln`. The VSTO host is a classic project that needs full MSBuild
  with the Office/VSTO workload, so the build goes through `Taskfile.yml`:
  `task compile`, `task test`, `task pack`. [Build and test](docs/src/content/docs/internals/build-and-test.md)
  lists the prerequisites, the targets and the five test passes.
- The facts about the virtual workbook are written once, in the abstract classes of
  `src/Allors.Excel.Tests`, and run against the headless implementation and against a real Excel.
  Add a behavioural test there so that both implementations are covered; add a test to
  `src/Allors.Excel.Interop.Tests` only for interop-only code such as `Cells.Chunks`.
- The interop test passes start a visible Excel for every test. They need an Excel that opens
  without a modal dialog, and they take minutes; the headless pass and the chunk tests run in
  seconds and need no Excel.
- Coordinates are zero-based in the library and one-based in Excel; keep the conversion at the
  interop boundary. A cell setter marks the cell dirty, and only `Flush` writes to Excel; keep
  that separation when adding a cell property, including its `AddDirty…` method, its dirty set
  and its `Render…` step.
- [Architecture](docs/src/content/docs/internals/architecture.md#known-defects) lists the known
  defects. Do not describe a defective member as working until the fix has shipped.
- Follow existing naming and structure; avoid new conventions without reason. Error messages
  should be actionable.

## Git

- Pull requests are the default. Creating a pull request requires approval.
- Use one branch and one pull request for an agreed body of work. Its parts land as focused
  commits on that branch. Create the pull request only when that work is complete; if one already
  exists, update it.
- Keep the title and description of a pull request correct for everything on its branch, and
  update them when the branch changes. For a squash merge, the title is the commit subject and the
  description is the commit body, describing the final changes and their validation.
- Use the conventional commit format `type(scope): description`. Keep commits focused and
  well-described. No AI attribution in commits or pull requests.
- Use purpose prefixes for branch names, such as `feature/`, `fix/`, `docs/` or `chore/`, not the
  name of an agent or tool.
- Nothing is published automatically. A release is a tag `vX.Y.Z`, which the `Release` workflow
  publishes to nuget.org; see [Build and test](docs/src/content/docs/internals/build-and-test.md#release).
  Changes limited to `docs/`, the documentation workflow and this file do not run `CI`.

## Workflow

- Verify that the build succeeds and the tests pass before considering a task complete:
  `task test`, or `task compile` and the headless and chunk passes when no usable Excel is
  available, saying so in the pull request.
- When fixing a bug, first write a failing test or amend an existing one. Add a new test to a
  suitable existing class. Do not change an existing test unless instructed to.
- Keep implementation facts beside the code: in tests, in code comments and in the documentation
  pages; do not maintain a second account of them here.

## Documentation

- `docs/src/content/docs/` is human documentation for users and maintainers. Before changing
  documentation, comments or public-facing explanations, read and apply the
  [writing guide](docs/src/content/docs/internals/writing-documentation.md) and the
  [terminology reference](docs/src/content/docs/terminology.md). The guide owns the audience,
  Diátaxis, status, Oxford English and visual explanation rules. Preserve exact identifiers and
  external names.
- Do not rename an established concept or substitute a competing term without prior agreement;
  explain the proposed meaning, reason and affected scope first. Apply an approved change
  consistently across code, prose, diagrams, examples and pull request text.
- Update pages made untrue by a change in the same pull request. Maintain the index, page
  statuses and links according to the writing guide. The README stays an introduction with
  absolute links to the pages, because it is packed into the NuGet packages.
- Date a verification that depends on an external state, such as a published package version or
  an observed failure, and say what was checked. Run a tutorial from a clean directory before
  changing its steps or its shown output.
- Before completing work, review the changed material for factual accuracy, agreed terminology,
  Oxford spelling and clear practitioner language, including labels, example output and pull
  request text. Editorial review is required; no prose linter is part of the tooling.
- The website in `docs/` has its own checks. Run them from `docs/` before completing documentation
  work, as [docs/README.md](docs/README.md) describes: `npm test`, `npm run check`,
  `npm run build`, `npm run check:links` and `npm run test:browser`.

## Build commands

The build uses [Task](https://taskfile.dev): `task compile`, `task test`, `task pack`, with
`CONFIGURATION=Release` for a release build. The documentation website in `docs/` builds with
npm; [docs/README.md](docs/README.md) lists its commands.
