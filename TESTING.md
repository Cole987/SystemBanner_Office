Testing on virtual machines
===========================

What the automated tests already cover, and what still needs a real Office on each version. Tick each line per VM.

Automated, on any Windows with the .NET Framework 3.5
------------------------------------------------------

- `build\SystemBannerOffice.Tests.exe`: the marking engine, built on the examples of the Manual. Expected: `41 passed, 0 failed`.
- `build\OfficeCheck.exe`: drives the installed Word, Excel and PowerPoint through automation (no install needed). Expected: `all checks passed`. Passed on Office 2016 (16.0), 64 bit, on Windows 10.

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
- [ ] `Simple` = 12 (France): Classify offers DIFFUSION RESTREINTE and SECRET only, and writes them in the banner.
- [ ] `OfficeSpellOut` = 1: banner lines spell control markings out.
- [ ] `OfficeRequireClassification` = 1: saving and printing an unclassified document asks to classify it first, in Word and Excel.

Word
- [ ] Classify an empty document SECRET//NOFORN: banner at the top and bottom of every page, red background, white text (Word 2000: text color may stay black, see README).
- [ ] A document with an existing header and page numbers keeps them under the banner.
- [ ] Several sections, first page different, odd and even pages: every page shows the banner, none twice.
- [ ] Portion mark on three paragraphs: (S//REL TO USA, GBR), (U), (TS//SI//NF). The banner becomes TOP SECRET//SI//NOFORN.
- [ ] Portion mark on an already marked paragraph replaces its mark.
- [ ] Authority block: written at the cursor, written again in place when changed.
- [ ] File > Properties > Custom: Classification and SystemBanner.* values are there.
- [ ] Save as .doc and reopen: banner, marks and properties are still there.

Excel
- [ ] Classify: the banner is in the header and footer of every sheet (print preview). Excel 2007 and later: colored text.

PowerPoint
- [ ] Classify or roll up: banners at the top and bottom of every slide, title slides included (PowerPoint 2000 to 2003 title master).
- [ ] Portion mark in a bullet, then Roll up.

Removal
- [ ] `Install\RemoveSystemBannerOffice.bat`: Office starts with no toolbar or tab and no error.

When Office does not start, or the add-in does not show
-------------------------------------------------------

1. Get Office back: start it in safe mode, which loads no add-in (`winword /safe`, `excel /safe`, `powerpnt /safe`; Office 2000 and XP have no safe mode switch, hold Ctrl while starting). If it still does not start, run `Install\RemoveSystemBannerOffice.bat` as an administrator.
2. Run `build\OfficeCheck.exe diag` as the user who started Office. It writes `SystemBannerOffice_diag.txt` on the Desktop: Office versions, the add-in's registration in Office and in COM, the add-ins Office disabled, a test creation of the add-in outside Office, and the end of the add-in's log.
3. The add-in's log is `%TEMP%\SystemBannerOffice.log`: one line per start (Office version, .NET version), toolbar or ribbon creation, and every error with its stack trace. No log at all means Office never started the add-in.

Send the diag file with the report.

Report
------

For each failure, note the VM, the Office version (Help > About), the button, the exact message, and attach the document if possible.
