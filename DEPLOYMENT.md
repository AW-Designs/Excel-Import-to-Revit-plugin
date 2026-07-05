# Deployment & Auto-Update

The add-in distributes through **GitHub Releases** and updates itself.

Repo: https://github.com/AW-Designs/Excel-Import-to-Revit-plugin

## How it works

```
You push a version tag  ─►  GitHub Actions builds R24/R25/R26  ─►  publishes a Release
                                                                          │
Colleague launches Revit  ─►  add-in checks the latest release  ◄─────────┘
                              (background, once every 4h max)
                                     │  newer version?
                                     ▼
                              downloads the matching zip, stages it,
                              swaps files in when Revit closes
                                     │
                              next launch runs the new version
```

- The running DLL is never overwritten while Revit holds it; the swap happens
  after Revit exits (a small detached process waits for it), so updates are safe.
- The updater downloads only the asset for that machine's Revit year
  (`ExcelScheduleImporter-R24/25/26.zip`).
- Assembly version is stamped from the tag, so version comparison is reliable.

## Cutting a release (you)

1. Commit your changes.
2. Bump and tag:
   ```
   git tag v1.0.1
   git push origin v1.0.1
   ```
3. GitHub Actions builds all three Revit versions and publishes the release
   automatically (watch the Actions tab). Nothing else to do.

Everyone on an installed copy gets it within one Revit restart (after the next
4-hour check window).

> Tag format is `vMAJOR.MINOR.PATCH` (e.g. `v1.2.0`). The leading `v` is stripped
> for the assembly version.

## First-time install (each colleague, once)

Run in PowerShell (no admin needed):

```powershell
irm https://github.com/AW-Designs/Excel-Import-to-Revit-plugin/releases/latest/download/install.ps1 -OutFile "$env:TEMP\esi-install.ps1"; & "$env:TEMP\esi-install.ps1"
```

It detects installed Revit versions, downloads the matching add-in, and registers
it under the user profile. On first launch, Revit may warn about an unsigned
add-in — click **Always Load** once.

## Local build (development)

`build.ps1` builds all three versions and deploys them to your own
`%AppData%\Autodesk\Revit\Addins\{2024,2025,2026}` for testing. This is separate
from the release pipeline.

## Notes

- **Code signing:** release builds from CI are unsigned, so colleagues click
  "Always Load" once. To remove that prompt org-wide, export the public
  certificate and have IT push it to the Trusted Publishers store via GPO.
- **Rate limits:** the updater checks at most once every 4 hours per machine to
  stay within GitHub's unauthenticated API limit (60/hour per IP).
- **Private repo:** if the repo is made private, the download URLs need a token;
  the current design assumes a public repo.
