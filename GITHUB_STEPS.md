# Build MR_Updater.exe on GitHub (no Python on your PC)

You only need a free GitHub account and a web browser. GitHub's own Windows
servers do the building. You end up with a single MR_Updater.exe to copy to
the clinic PC.

## The two files you upload
- MR_Updater.py            -> goes at the repo ROOT
- build-windows.yml        -> goes at  .github/workflows/build-windows.yml

================================================================
STEP BY STEP
================================================================

1) Create the repository
   - Go to github.com, sign in (or sign up - free).
   - Click the "+" (top right) -> New repository.
   - Name it e.g.  mr-updater  -> set it Private -> Create repository.

2) Upload MR_Updater.py
   - On the new repo page click  "uploading an existing file"
     (or  Add file -> Upload files).
   - Drag MR_Updater.py into the page.
   - Click  Commit changes.

3) Add the workflow file (this creates the folders for you)
   - Click  Add file -> Create new file.
   - In the file-name box type EXACTLY:
         .github/workflows/build-windows.yml
     (typing the slashes auto-creates the folders)
   - Open build-windows.yml that I gave you, copy ALL of it,
     paste it into the big text area.
   - Click  Commit changes.

   * Tip: as soon as you commit this, the build may start automatically.

4) Run / watch the build
   - Click the  Actions  tab (top of the repo).
   - Open the run called "Build Windows EXE".
     If it didn't auto-start: click it, then "Run workflow" -> Run.
   - Wait ~2-3 minutes until it shows a green check mark.

5) Download the EXE
   - Still on that finished run, scroll down to  Artifacts.
   - Click  MR_Updater-exe  -> it downloads a .zip.
   - Unzip it -> inside is  MR_Updater.exe.

6) Use it
   - Copy MR_Updater.exe to the clinic PC (USB stick or email).
   - Double-click it. No Python, no install, no admin needed.
   - Browse for the report, type the new lab date (M/D/YYYY), click Generate.

================================================================
IF SOMETHING GOES WRONG
================================================================
- Red X on the build? Open the failed step, copy the red error text, send it
  to me and I'll fix the workflow.
- Windows SmartScreen may warn the first time you open the .exe
  ("unknown publisher"): click  More info -> Run anyway. That's normal for a
  self-built app; it is not blocked, just unsigned.
- The clinic blocks the download? Do steps 1-5 at home, put the .exe on a USB.
