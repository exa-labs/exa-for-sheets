#!/usr/bin/env node
/**
 * Version bump script for Exa for Google Sheets
 * Syncs version across package.json, Code.gs, and CHANGELOG.md
 * Optionally commits, tags, and pushes the release
 * 
 * Usage: node scripts/bump-version.js <major|minor|patch> [--tag]
 *   --tag: Also commit, create git tag, and push
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const ROOT = path.join(__dirname, '..');

function parseVersion(version) {
  const match = version.match(/^(\d+)\.(\d+)\.(\d+)$/);
  if (!match) throw new Error(`Invalid version: ${version}`);
  return { major: parseInt(match[1]), minor: parseInt(match[2]), patch: parseInt(match[3]) };
}

function bumpVersion(current, type) {
  const v = parseVersion(current);
  switch (type) {
    case 'major': return `${v.major + 1}.0.0`;
    case 'minor': return `${v.major}.${v.minor + 1}.0`;
    case 'patch': return `${v.major}.${v.minor}.${v.patch + 1}`;
    default: throw new Error(`Invalid bump type: ${type}. Use major, minor, or patch.`);
  }
}

function updatePackageJson(newVersion) {
  const filePath = path.join(ROOT, 'package.json');
  const content = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  const oldVersion = content.version;
  content.version = newVersion;
  fs.writeFileSync(filePath, JSON.stringify(content, null, 2) + '\n');
  return oldVersion;
}

function updateCodeGs(newVersion) {
  const filePath = path.join(ROOT, 'Code.gs');
  let content = fs.readFileSync(filePath, 'utf8');
  content = content.replace(/Version: \d+\.\d+\.\d+/, `Version: ${newVersion}`);
  fs.writeFileSync(filePath, content);
}

function updateChangelog(newVersion) {
  const filePath = path.join(ROOT, 'CHANGELOG.md');
  let content = fs.readFileSync(filePath, 'utf8');
  const today = new Date().toISOString().split('T')[0];
  content = content.replace(
    /## \[\d+\.\d+\.\d+\] - \d{4}-\d{2}-\d{2}/,
    `## [${newVersion}] - ${today}`
  );
  fs.writeFileSync(filePath, content);
}

function gitTag(version) {
  const tagName = `v${version}`;
  try {
    execSync('git add package.json Code.gs CHANGELOG.md', { cwd: ROOT, stdio: 'inherit' });
    execSync(`git commit -m "chore: bump version to ${version}"`, { cwd: ROOT, stdio: 'inherit' });
    execSync(`git tag ${tagName}`, { cwd: ROOT, stdio: 'inherit' });
    console.log(`\nCreated tag: ${tagName}`);
    console.log('Run: git push && git push --tags');
  } catch (err) {
    console.error('Git operation failed:', err.message);
    process.exit(1);
  }
}

function main() {
  const args = process.argv.slice(2);
  const bumpType = args.find(a => ['major', 'minor', 'patch'].includes(a));
  const shouldTag = args.includes('--tag');

  if (!bumpType) {
    console.error('Usage: node scripts/bump-version.js <major|minor|patch> [--tag]');
    console.error('  --tag: Also commit and create git tag');
    process.exit(1);
  }

  const packageJson = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  const currentVersion = packageJson.version;
  const newVersion = bumpVersion(currentVersion, bumpType);

  console.log(`Bumping version: ${currentVersion} -> ${newVersion}`);

  updatePackageJson(newVersion);
  console.log('  Updated package.json');

  updateCodeGs(newVersion);
  console.log('  Updated Code.gs');

  updateChangelog(newVersion);
  console.log('  Updated CHANGELOG.md');

  console.log(`\nVersion bumped to ${newVersion}`);

  if (shouldTag) {
    gitTag(newVersion);
  } else {
    console.log('Run: git add -A && git commit -m "chore: bump version to ' + newVersion + '"');
    console.log('Or use --tag flag to auto-commit and tag');
  }
}

main();
