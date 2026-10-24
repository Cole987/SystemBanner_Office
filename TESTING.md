Testing on virtual machines
===========================

What the automated tests already cover, and what still needs a real Office on each version. Tick each line per VM.

Automated, on any Windows with the .NET Framework 3.5
------------------------------------------------------

- `build\SystemBannerOffice.Tests.exe`: the marking engine, built on the examples of the Manual. Expected: `41 passed, 0 failed`.
- `build\OfficeCheck.exe`: drives the installed Word, Excel and PowerPoint through automation (no install needed). Expected: `all checks passed`. Passed on Office 2016 (16.0), 64 bit, on Windows 10.

- `build\OfficeCheck.exe toolbar`: starts the add-in in Word, Excel and PowerPoint as Office does, with the Office 2000-2003 toolbar on any version; checks its 5 buttons, that a hidden or deleted toolbar comes back, and that it goes away when the add-in unloads. Expected: `all checks passed`.

- `build\OfficeCheck.exe dialogs`: opens each dialog filled with a marking, checked list boxes included (also part of the plain `OfficeCheck.exe` run). Expected: `all checks passed`.

Run both first on each VM: they show at once whether the late bound calls work on that Office version.

Matrix
------

| VM | Windows | Office | Bitness | Tests.exe | OfficeCheck.exe | Manual checks |
|---|---|---|---|---|---|---|
| 1 | 2000 SP4 | 2000 | 32 | | | |
| 2 | XP SP3 | XP or 2003 | 32 | | | |
| 3 | Vista SP2 or 7 | 2007 | 32 | | | |
| 4 | 7 x64 | 2010 | 32 | | | |
| 5 | 7 x64 | 2010 | 64 | | | |

Manual checks, after `Install\InstallSystemBannerOffice.bat` as an administrator
------------------------------------------------------------------------------

Loading
- [ ] Word, Excel and PowerPoint start with no error message.
- [ ] Office 2000 to 2003: the SystemBanner toolbar shows its four buttons. Office 2007 and 2010: the Classification tab shows them.
- [ ] Closing and reopening Office leaves a single toolbar (Office 2000 to 2003).

Settings (`HKLM\SOFTWARE\Policies\SystemBanner`, or the SystemBanner Group Policy)
- [ ] `Simple` = 4 (SECRET): Classify refuses TOP SECRET, with the reason in the dialog.
- [ ] `Simple` = 6: SCI accepted. `Simple` = 5: SCI refused.
- [ ] `Simple` = 12 (France SECRET): Classify goes straight to the French dialog, offers Diffusion Restreinte and Secret, refuses Très Secret with the reason.
- [ ] `Simple` = 4 (United States): Classify goes straight to the IC dialog, no French markings.
- [ ] No `Simple` value: Classify asks for the system: France, Germany, United Kingdom or IC.
- [ ] `Simple` = 23 (Germany GEHEIM): Classify goes straight to the German dialog and refuses STRENG GEHEIM.
- [ ] `Simple` = 32 (United Kingdom OFFICIAL-SENSITIVE): Classify goes straight to the British dialog and refuses SECRET.

Language
- [ ] An English Office shows the buttons and dialogs in English, a French Office in French.
- [ ] `OfficeLanguage` = `fr` (or `en`) overrides the language of Office.
- [ ] `OfficeSpellOut` = 1: banner lines spell control markings out.
- [ ] `OfficeRequireClassification` = 1: saving and printing an unclassified document asks to classify it first, in Word and Excel.

Word
- [ ] Classify an empty document SECRET//NOFORN: banner at the top and bottom of every page, in bold red text with no background (Word 2000: text color may stay black, see README); printed in black and white, it stays readable. `OfficeBannerBackground` = 1: white text on a red background.
- [ ] A document with an existing header and page numbers keeps them under the banner.
- [ ] Several sections, first page different, odd and even pages: every page shows the banner, none twice.
- [ ] Portion mark on three paragraphs: (S//REL TO USA, GBR), (U), (TS//SI//NF). The banner becomes TOP SECRET//SI//NOFORN.
- [ ] Portion mark on an already marked paragraph replaces its mark.
- [ ] Authority block: written at the cursor, written again in place when changed.
- [ ] File > Properties > Custom: Classification and SystemBanner.* values are there.
- [ ] Save as .doc and reopen: banner, marks and properties are still there.

French markings (Word)
- [ ] Classify as Secret, Spécial France: SECRET in red, framed, at the top and bottom of every page, SPÉCIAL FRANCE in blue at the top only, "page X / Y" under the bottom stamp.
- [ ] The identification dialog opens: with a deadline more than 50 years away, or a re-evaluation more than 20 years away, Write stays greyed with the reason.
- [ ] Portion mark on three paragraphs: [NP], [S SF], [DR]. Roll up: the document is SECRET SPÉCIAL FRANCE.
- [ ] Diffusion Restreinte: stamp at the top only, no bottom stamp and no page numbers.
- [ ] Très Secret: copy number and reference required; the reference shows before the page number on every page.
- [ ] Bound document: the cover stamp with the penal code warning at the bottom of the first page only.
- [ ] Downgrade: Très Secret to Secret changes the stamps and writes the annexe 38 stamp; Declassify removes the stamps.
- [ ] Switching the same document to IC markings replaces the stamps with the banner line, with no frame left.
- [ ] Word 2000: the stamps are red (or the nearest Word 2000 color) and framed.

German markings (Word)
- [ ] GEHEIM: the level and "amtlich geheim gehalten" in red at the top and bottom of every page, "Seite X von Y" under the bottom stamp.
- [ ] VS-VERTRAULICH: at the top only, in blue; VS-NfD: the level alone at the top, in black.
- [ ] Part marks [offen], [VS-NfD], [Geh.]: Roll up gives GEHEIM.
- [ ] A checked Sperrvermerk shows under the top stamp.
- [ ] First page: from VS-VERTRAULICH up, the reference, the copy and recipient, and the end year (30 years at most) are required.
- [ ] Downgrade: GEHEIM to VS-VERTRAULICH changes the stamps and writes the note; lifting removes them.

British markings (Word)
- [ ] OFFICIAL-SENSITIVE with RECIPIENTS ONLY and PERSONAL DATA: the marking at the top and bottom of every page, in black.
- [ ] RECIPIENTS ONLY is refused on OFFICIAL, national caveats on OFFICIAL, FOR PUBLIC RELEASE with -SENSITIVE.
- [ ] Portion mark explains that the GSCP has no portion marking.

EU markings (Word)
- [ ] Classify as SECRET UE/EU SECRET with a reference, copy number 2 and RELEASABLE TO NATO: the level at the top and bottom, the additional marking under the top one, "reference - Ex./Copy 2 - 1 / 1" at the bottom.
- [ ] CONFIDENTIEL UE/EU CONFIDENTIAL: no copy number at the bottom; no reference: Classify stays greyed with the reason.
- [ ] Paragraph marks [R-UE/EU-R] and [C-UE/EU-C], Roll up: CONFIDENTIEL UE/EU CONFIDENTIAL; the reference and additional markings stay.
- [ ] Downgrade asks for the originator's consent, then opens the EU dialog.
- [ ] `Simple` = 62: Classify goes straight to the EU dialog and refuses SECRET UE/EU SECRET.

Excel
- [ ] Classify: the banner is in the header and footer of every sheet (print preview). Excel 2007 and later: colored text.

PowerPoint
- [ ] Classify or roll up: banners at the top and bottom of every slide, title slides included (PowerPoint 2000 to 2003 title master).
- [ ] Portion mark in a bullet, then Roll up.

Setup program
- [ ] `Install\SystemBannerOfficeSetup.exe`, Install: the add-in loads in Word, Excel and PowerPoint; "SystemBanner Office add-in" is in Add or Remove Programs.
- [ ] Windows Vista and later: `SystemBannerOffice.admx` and its ADML are in `C:\Windows\PolicyDefinitions`; gpedit.msc shows *SystemBanner Office add-in* with 6 settings. Windows 2000 and XP: the ADM files are in `C:\WINDOWS\inf`, Add/Remove Templates loads them.
- [ ] Each setting set in gpedit.msc writes its value in `HKLM\SOFTWARE\Policies\SystemBanner` and the add-in follows it after Office restarts.
- [ ] With Word open, Install refuses and names it. As a standard user (Windows 2000 and XP), Install refuses.
- [ ] Office 2000 to 2003: `WINWORD.EXE.config`, `EXCEL.EXE.config`, `POWERPNT.EXE.config` written next to Office.
- [ ] `/install /quiet` and `/uninstall /quiet` from a command prompt: no window, `echo %errorlevel%` gives 0.
- [ ] Remove (or Add or Remove Programs): Office starts with no toolbar or tab and no error; the folder, the templates and the config files are gone (after a restart if Office held a file).

Removal
- [ ] `Install\RemoveSystemBannerOffice.bat`: Office starts with no toolbar or tab and no error.

Excel 2000 to 2003
- [ ] With no printer installed, Classify explains that Excel needs one; with the "Generic / Text Only" printer, the header and footer are written.

When Office does not start, or the add-in does not show
-------------------------------------------------------

1. Get Office back: start it in safe mode, which loads no add-in (`winword /safe`, `excel /safe`, `powerpnt /safe`; Office 2000 and XP have no safe mode switch, hold Ctrl while starting). If it still does not start, run `Install\RemoveSystemBannerOffice.bat` as an administrator.
2. Run `build\OfficeCheck.exe load word` (or `excel`, `powerpoint`): Office loads the installed add-in as the COM add-ins dialog does, and the end of the log shows how far it went.
3. Run `build\OfficeCheck.exe diag` as the user who started Office. It writes `SystemBannerOffice_diag.txt` on the Desktop: Office versions, the add-in's registration in Office and in COM, the add-ins Office disabled, a test creation of the add-in outside Office, and the end of the add-in's log.
4. The log tells how far each click went: `OnAction Classify` (Office 2007 and later called the add-in), then `command Classify, dialogs owned by window N` (the command started; N = 0 means no Office window was found for the dialog). No `OnAction` line after a click on the ribbon: in Office 2007, tick Word Options > Advanced > General > Show add-in user interface errors, and note the message. Toolbar: `toolbar created with 5 buttons`, then up to 3 `toolbar check` lines in the 6 seconds after startup.
5. In Word or Excel 2000 to 2003, `File or assembly name SystemBannerOffice, or one of its dependencies, was not found` (from `OfficeCheck.exe load word`) is a .NET 1.x message: Office started .NET 1.1, which cannot load the add-in. The installer writes `WINWORD.EXE.config`, `EXCEL.EXE.config` and `POWERPNT.EXE.config` next to Office 2000 to 2003 to ask for .NET 2.0; the diag shows them under each Office version. An existing config file is never replaced: the installer warns, and `supportedRuntime version="v2.0.50727"` must be added to it by hand.
6. The add-in's log is `%TEMP%\SystemBannerOffice.log`: one line per start (Office version, .NET version), toolbar or ribbon creation, and every error with its stack trace. No log at all means Office never started the add-in.

Send the diag file with the report.

Report
------

For each failure, note the VM, the Office version (Help > About), the button, the exact message, and attach the document if possible.
