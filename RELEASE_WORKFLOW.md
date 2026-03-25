# Release Workflow

## Same-version fix

Use this when the version number should stay the same and the existing GitHub Release assets should be replaced.

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode fix
```

## New version release

Use this when you want to bump the patch version automatically.

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode release
```

## New version release with explicit version

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode release -Version 1.1.4
```

## What the script does

1. Reads `version.txt`
2. Keeps the version as-is for `fix`, or bumps patch for `release`
3. Syncs installer version metadata
4. Builds the installer and zip package
5. Commits and pushes source changes to GitHub
6. Updates the existing GitHub Release if the tag already exists
7. Creates a new GitHub Release if the tag does not exist
8. Generates a humorous release note markdown file in `dist\`
