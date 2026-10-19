import versionJson from '../version.json' with { type: 'json' };

// The documented version follows version.json, which Nerdbank.GitVersioning stamps on the packages.
// Releases are built from tags, so the packages built from main are unreleased and the label says so.
export const title = 'Allors Excel';
export const description = 'Documentation for developers who build Excel add-ins with Allors Excel, and for its maintainers.';
export const version = `main · ${versionJson.version} (unreleased)`;
export const repository = 'https://github.com/allors/excel';
export const branch = 'main';
export const site = process.env.DOCS_SITE || 'https://allors.github.io';
const trimmed = (process.env.DOCS_BASE || '/excel/').split('/').filter(Boolean).join('/');
export const base = trimmed ? `/${trimmed}/` : '/';
