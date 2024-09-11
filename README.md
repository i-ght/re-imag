# Re-imaging Tool

Automates completing re-imaging process of corporate windows computers.

What needs to be done for re-imaging to be considered complete:

1. Install new OS from USB stick
2. Copy shortcut to legacy apps onto public desktop.
3. Install chrome.
4. - Install drivers software depending on computer manufacturer.
   - If drivers program installs, run it and install latest drivers; ignore bios updates.
5. Rename the computer to what it was named previously (before re-imaging process, often labeled on hardware).
6. Add printers for every window user.
7. Add Scanners, other peripherals.

What the script does:
1. Mounts a network drive containing necessary files.
2. Copies shortcut to legacy apps onto public desktop.
3. Launches chrome setup and awaits user to complete the setup. (If exit code is not 0, user is warned.)
4. Launches driver setup and awaits; if setup installs, runs the installed program.
5. Renames the computer to user specification.
6. Adds specified printers.