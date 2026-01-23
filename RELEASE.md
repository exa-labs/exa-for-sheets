# Release Process

This document outlines the steps to release a new version of Exa for Sheets.

## Prerequisites

- Ensure you're logged in to clasp: `npm run login`
- Ensure all changes are committed to git
- Ensure you have push access to the repository

## Automated Release Steps

### Using npm version command

1. **Bump version and create git tag**
   ```bash
   npm version patch   # For bug fixes (1.0.1 → 1.0.2)
   npm version minor   # For new features (1.0.1 → 1.1.0)
   npm version major   # For breaking changes (1.0.1 → 2.0.0)
   ```
   This automatically:
   - Updates version in `package.json`
   - Creates a git commit with message "1.0.2"
   - Creates a git tag (e.g., "1.0.2")

2. **Push code to Google Apps Script**
   ```bash
   npm run push
   ```

3. **Create an immutable version in Apps Script**
   ```bash
   npm run version "Release v1.0.2"
   ```
   Replace version number with your current version. This creates a snapshot in Apps Script.

4. **Create a deployment**
   ```bash
   npm run deploy
   ```
   Or with specific version and description:
   ```bash
   clasp deploy -V <versionNumber> -d "Release v1.0.2"
   ```

5. **Push to GitHub**
   ```bash
   git push origin master --tags
   ```

## Manual Release Steps

If you prefer manual control:

1. **Update version in `package.json`**
   - Change the version number following semver

2. **Commit the version change**
   ```bash
   git add package.json
   git commit -m "Bump version to 1.0.2"
   git tag 1.0.2
   ```

3. **Push to Apps Script**
   ```bash
   npm run push
   ```

4. **Create Apps Script version**
   ```bash
   npm run version "Release v1.0.2"
   ```

5. **Deploy**
   ```bash
   npm run deploy
   ```

6. **Push to GitHub**
   ```bash
   git push origin master --tags
   ```

## Available clasp Commands

- `npm run version` - Create immutable version snapshot
- `npm run versions` - List all versions
- `npm run deploy` - Create new deployment
- `clasp deployments` - List all deployments
- `clasp undeploy <deploymentId>` - Remove a deployment

## Notes

- Custom functions may take several minutes to update in Google Sheets after deployment
- Users may need to close and reopen spreadsheets to see updates
- Deployments are separate from versions - versions are snapshots, deployments are what users install
- Check deployment status: `clasp deployments`
