# NewUserAutomation

Desktop tool for new user onboarding workflows (Graph, Exchange, SharePoint), with customer profile support and guided live-run execution.

## What this repo includes

- WPF desktop app source (`NewUserAutomation.App`)
- Core libraries and test projects
- Installer script (`scripts/Install-NewUserAutomation.ps1`)
- Installer EXE source (`scripts/InstallerLauncher`)
- Release build script (`scripts/publish.ps1`)
- GitHub Actions release workflow (`.github/workflows/release.yml`)

## Install for end users (EXE installer)

1. Download `NewUserAutomationInstaller.exe`.
2. Run the EXE (double-click or right-click Run as administrator if needed).
3. The installer calls GitHub Releases `latest` for this repo.
4. It downloads and installs/updates into `%LOCALAPPDATA%\NewUserAutomation\current`.
5. Launch `NewUserAutomation` from Start Menu or Desktop shortcut.

### GitHub release lookup

The EXE installer now reads:

- `https://api.github.com/repos/The-Emerald-Group/newuserautomation/releases/latest`

It downloads the `NewUserAutomation-win-x64.zip` asset from the latest release and compares installed version against the latest tag version.

## Install for end users (PowerShell script fallback)

Use the installer script:

1. Download `scripts/Install-NewUserAutomation.ps1`.
2. Right-click the script and choose **Run with PowerShell**.
3. Follow prompts (it self-elevates if admin rights are needed).
4. Launch `NewUserAutomation` from Start Menu or Desktop shortcut.

The installer:
- Downloads the latest configured app zip
- Extracts and installs to `%LOCALAPPDATA%\NewUserAutomation\current` (default)
- Stops running app instances before replacing files
- Creates Start Menu + optional Desktop shortcuts

## Updating deployed installs

- Preferred: users re-run `NewUserAutomationInstaller.exe`; it checks GitHub latest release tag/version and updates.
- Fallback: re-run `Install-NewUserAutomation.ps1`; it downloads from the configured zip URL.
- Customer data/certs are expected to remain outside packaged app output.

## Release process (maintainers)

1. Commit and push changes to `main`.
2. Create and push a version tag:
   - `git tag v1.0.0`
   - `git push newuserautomation v1.0.0`
3. GitHub Actions will:
   - Build a self-contained `win-x64` publish
   - Create `dist/NewUserAutomation-win-x64.zip`
   - Build `NewUserAutomationInstaller.exe`
   - Attach both files to the GitHub Release for that tag

You can also trigger the workflow manually from the Actions tab (`Build and Release`).

4. Distribute `NewUserAutomationInstaller.exe` to users.
5. For future updates, publish a new release tag with updated assets; users can rerun the same installer EXE.

## Configure installer download source

The installer uses `scripts/Install-NewUserAutomation.ps1` `$DefaultZipUrl`.

Set this URL to your preferred hosted release package, for example:
- Your internal web host
- A GitHub Releases direct asset URL (only practical when repo/assets are publicly accessible)

## Repo visibility note

This GitHub-driven installer flow requires release assets to be publicly downloadable (or authenticated if private).
Since your repo is public, no client-side token is needed for installer updates.

## Local build and package

From repo root:

- Build app:
  - `dotnet build .\NewUserAutomation.App\NewUserAutomation.App.csproj -c Release`
- Publish zip:
  - `.\scripts\publish.ps1 -Runtime win-x64 -Configuration Release -ZipOutput`

## Notes

- Keep secrets, certificates, and local customer data out of git (covered by `.gitignore`).
- Do not commit generated installer binaries (`scripts/InstallerLauncher.exe` / `.pdb`).
