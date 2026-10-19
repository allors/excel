import versionJson from '../version.json' with { type: 'json' };

// The documented version follows version.json, which Nerdbank.GitVersioning stamps on the packages.
// Releases are built from tags, so the packages built from main are unreleased and the label says so.
export const title = 'Allors Excel';
export const description = 'Documentation for developers who build Excel add-ins with Allors Excel, and for its maintainers.';
export const version = `main · ${versionJson.version} (unreleased)`;
export const repository = 'https://github.com/allors/excel';
export const branch = 'main';
export const site = process.env.DOCS_SITE || 'https://docs.allors.com';
const trimmed = (process.env.DOCS_BASE || '/excel/').split('/').filter(Boolean).join('/');
export const base = trimmed ? `/${trimmed}/` : '/';
// The documentation home is published at the root of the same origin by allors/allors.github.io. The
// sidebar links to it, and the link checker leaves that link to the home's own build.
export const home = new URL('/', site).href;
