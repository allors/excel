# AGENTS.md

Instructions for agents working in this repository. The technical facts live in the human
documentation; this file says where to find them and how to work here.

## Change policy

Visibility: public. Change policy: conservative.

- Make only changes required by the explicit request. Preserve existing architecture,
  organization and conventions unless directly instructed otherwise.
- Follow existing test placement, fixtures and patterns; do not introduce a separate
  testing approach. Preserve established test expectations unless the requested change
  explicitly changes that behaviour.
- Before adding test coverage, inspect the closest existing test methods. When a method
  already checks the same behaviour across types or cases, extend it in place. Add a
  separate method only for a distinct scenario that does not fit the existing test's purpose.
- Keep unrelated test setup, timing and expectations unchanged. Investigate validation
  failures and fix regressions caused by the requested change. Report independent failures
  and resulting validation limits separately; passing validation does not authorize
  unrelated fixes.
- Before committing, review the test diff against these rules. Check every new test method
  and every change to existing setup or expectations for necessity within the requested scope.
- Edit documentation, including README files, changelogs, diagrams and guides, only when
  explicitly requested. Report related improvements or documentation implications
  without implementing them.
- An exception applies only to the change directly authorized; it does not waive the
  remaining rules. Change this policy only when explicitly instructed.

## Research records and agent memory

- Store agent research notes, external comparisons, code-evidence inventories and
  investigation checkpoints in the agent's own built-in memory outside Git. Do not create
  committed memory files or a separate shared memory store for these records.
- Do not append sections such as "Research basis and code evidence" merely to record how
  the agent investigated a task. Keep resulting guidance, designs, decisions and essential
  rationale in repository documentation when that documentation is authorized.
- Retain supporting references where a substantive documented claim depends on them.
  Include an inspected revision or date only when needed to interpret or verify that claim;
  routine investigation stamps belong in memory.
- Keep repository documentation understandable without access to any agent's memory.
  Memory is working context, not authoritative product documentation; verify remembered
  implementation details against the source relevant to the task.
- Keep agent-only handoff notes in built-in memory and provide the necessary continuation
  context in chat. If memory is unavailable, use chat and state the limitation; do not claim
  to have saved memory or substitute committed files. Explicitly requested research
  deliverables and project plans remain governed by their authorized scope.

## Working agreement

- Stay within the approved objective and preserve unrelated work. Keep explanations and
  authorized documentation concise.
- Before asking for a decision, explain it fully in chat: the mechanism and relevant code,
  consequences, options, actual costs and available measurements, and a recommendation
  including the argument against it. Ask the question in plain text at the end.
- For long work explicitly divided into subtasks, describe the current subtask's changes,
  files or systems and validation before starting. Approval of a multi-subtask plan
  authorizes only its first subtask.
- At each boundary, summarize changes and verification, describe the next subtask and its
  validation, then wait for explicit approval. Proceed across boundaries only when the
  user explicitly authorizes it for the session. Re-check materially expanded scope.
- Do not invent subtask boundaries for simple tasks or routine implementation steps.
- At agreed boundaries, assess whether a fresh session would help. Before recommending
  one, record the branch, commit SHAs, validation, uncommitted work, pending decisions and
  exact next step. Use explicitly authorized project notes for project deliverables; keep
  agent-only notes in built-in memory and provide the necessary continuation context in chat.
  Include a 2–4-line restart prompt; do not require a restart after every small task.

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
- Error messages should be actionable.

## Git

- Pull requests are the default. Creating or merging a pull request requires explicit
  approval. Keep one branch and one PR for an agreed body of work, with focused commits;
  open the PR when that work is complete.
- Use purpose prefixes such as `feature/`, `fix/`, `docs/` or `chore/`, not agent/tool names.
  Use conventional commits (`type(scope): description`) with no AI attribution.
- Keep the PR title and body accurate for the final diff. Update them when the branch's
  scope changes, including after review commits; describe the resulting behaviour and
  validation, not the development history.
- Default to squash merge. Use the PR title as the commit subject and its full description
  as the body, with wording that remains true after the merge.

## Release

- Nothing is published automatically. A release is a tag `vX.Y.Z`, which the `Release` workflow
  publishes to nuget.org; see [Build and test](docs/src/content/docs/internals/build-and-test.md#release).
  Changes limited to `docs/`, the documentation workflow and this file do not run `CI`.

## Workflow

- For code changes, verify that the build succeeds and the tests pass before considering the
  task complete: `task test`, or `task compile` and the headless and chunk passes when no usable
  Excel is available, saying so in the pull request.
- Follow test-driven development for requested behaviour changes. Add a failing regression case
  in a suitable existing class before implementing a fix.
- Use tests, code comments and human documentation for implementation facts; do not maintain
  a second account of them here.

## Documentation

- `docs/src/content/docs/` is human documentation for users and maintainers. For explicitly
  requested documentation work, read and apply the
  [writing guide](docs/src/content/docs/internals/writing-documentation.md) and the
  [terminology reference](docs/src/content/docs/terminology.md). The guide owns the audience,
  Diátaxis, status, Oxford English and visual explanation rules. Preserve exact identifiers and
  external names.
- Do not rename an established concept or substitute a competing term without prior agreement;
  explain the proposed meaning, reason and affected scope first. Apply an approved change
  consistently within the explicitly authorized scope; report implications outside it.
- Within authorized documentation work, maintain the affected index, page statuses and links
  according to the writing guide. The README stays an introduction with absolute links to the
  pages, because it is packed into the NuGet packages.
- In authorized documentation, date a verification that depends on an external state, such as
  a published package version or an observed failure, and say what was checked. Run a tutorial
  from a clean directory before changing its steps or its shown output.
- Before completing authorized documentation work, review the changed material for factual
  accuracy, agreed terminology, Oxford spelling and clear practitioner language, including
  labels, example output and pull request text. Editorial review is required; no prose linter
  is part of the tooling.
- For explicitly requested website documentation changes, run the checks from `docs/`, as
  [docs/README.md](docs/README.md) describes: `npm test`, `npm run check`,
  `npm run build`, `npm run check:links` and `npm run test:browser`.

## Build commands

The build uses [Task](https://taskfile.dev): `task compile`, `task test`, `task pack`, with
`CONFIGURATION=Release` for a release build. The documentation website in `docs/` builds with
npm; [docs/README.md](docs/README.md) lists its commands.
