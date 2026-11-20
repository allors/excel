import versionJson from '../version.json' with { type: 'json' };

// The documented version follows version.json, which Nerdbank.GitVersioning stamps on the packages.
// Releases are built from tags, so the packages built from main are unreleased and the label says so.
export const title = 'Allors Excel';
export const description = 'Documentation for developers who build Excel add-ins with Allors Excel, and for its maintainers.';
export const version = `main · ${versionJson.version} (unreleased)`;
export const repository = 'https://github.com/allors/excel';
export const branch = 'main';
// The root of the repository, which holds this website in docs/. A link from a page to a file that is not
// a page becomes a link to that file on GitHub, by its path from this root.
export const repositoryRoot = new URL('../', import.meta.url);
// The audiences that a page can name, in the order in which the agent index lists them. The content
// schema accepts these names and no others.
export const audiences = ['Users and maintainers', 'Users', 'Maintainers'];
export const site = process.env.DOCS_SITE || 'https://docs.allors.com';
const trimmed = (process.env.DOCS_BASE || '/excel/').split('/').filter(Boolean).join('/');
export const base = trimmed ? `/${trimmed}/` : '/';
// The documentation home is published at the root of the same origin by allors/allors.github.io. The
// sidebar links to it, and the link checker leaves that link to the home's own build.
export const home = new URL('/', site).href;
