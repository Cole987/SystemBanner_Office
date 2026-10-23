SystemBanner Office
===================

A Microsoft Office add-in that marks documents with their classification, from the same settings as [SystemBanner](https://github.com/Cole987/SystemBanner_XP): banner lines at the top and bottom of every page, portion marks, roll-up of the portion marks into the banner line, and the classification authority block. One build loads in Office 2000 to 2010 (Word, Excel, PowerPoint), like the Windows versions of SystemBanner.

It follows the *Intelligence Community Markings System Register and Manual* (ODNI, effective 31 December 2016, released under FOIA in 2021) for United States markings, and the stock classifications of SystemBanner for France, Germany, the United Kingdom, NATO and the United Nations.

Like SystemBanner, this is a retro computing and preservation project, shared as is. It has not been evaluated or certified for computers that handle real classified or sensitive information.

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Settings](#settings)
- [Use](#use)
- [What the add-in checks](#what-the-add-in-checks)
- [Limits](#limits)
- [Building and testing](#building-and-testing)

--------
FEATURES
--------

- **Banner lines.** Top and bottom of every page in Word (as their own paragraph in the header and footer, the rest of the header stays), every sheet in Excel, and the slide master in PowerPoint, with the colors of the SystemBanner presets.
- **The whole IC Register.** US classifications, non-US (`//GBR SECRET`), NATO (`//COSMIC TOP SECRET`, NAC activities) and JOINT (`//JOINT SECRET CAN GBR USA`) markings, SCI control systems with compartments and sub-compartments, SAP programs, Atomic Energy Act markings (RD, CNWDI, SIGMA, FRD, DOD UCNI, DOE UCNI, TFNI), FGI, the 13 dissemination controls and the 8 non-IC dissemination controls, in Register order with its separators.
- **Portion marks.** "Portion mark" puts `(S//REL TO USA, GBR)` at the start of the paragraph at the cursor, replacing the mark already there, in Word and PowerPoint.
- **Roll-up.** The banner line is rebuilt from the portion marks: highest classification, every unique SCI, SAP and AEA value, FGI sources, the 27 FD&R precedence rules of Table 3 (NOFORN, REL TO list intersections with FVEY, ACGU and TEYE expansion, RELIDO, EYES ONLY, DISPLAY ONLY), and the supersession rules of each marking template (RD over FRD and TFNI, ORCON over ORCON-USGOV, NODIS over EXDIS, FOUO dropped from classified banners, SBU NOFORN and LES NOFORN...).
- **Classification authority block.** Classified By, Reason or Derived From, Declassify On, checked as it is typed: date format and 25 year limit, exemptions (`25X#`, `50X1-HUM`, `50X2-WMD`, `75X#`), retired values (OADR, MR, X1 to X8, `25X1-human`) with what to write instead, and the `N/A to RD/FRD/TFNI and NATO portions` line. The most restrictive Declassify On of several sources follows the hierarchy of section E.3.
- **Metadata.** The classification is stored in the document properties (`Classification`, `SystemBanner.Banner`, the authority block), for mail guards and file servers.
- **The computer's classification is the ceiling.** A document cannot be marked above the `Simple` level of the computer; SCI needs a TOP SECRET//SCI computer. On a computer set to the French, German, British, NATO or UN system, the add-in offers that system's levels instead of IC markings.
- **No saving or printing unclassified documents**, when the administrator asks for it (Word and Excel).

------------
REQUIREMENTS
------------

- Windows 2000 SP4 to Windows 7 (and later), with the .NET Framework 2.0 or later (3.5 SP1 on Windows 7 and later).
- Office 2000, XP, 2003, 2007 or 2010, 32 bit, or 64 bit Office 2010. The ribbon tab appears in Office 2007 and later, a toolbar in Office 2000 to 2003.

------------
INSTALLATION
------------

As an administrator, run `Install\InstallSystemBannerOffice.bat`. It copies the add-in to `C:\Program Files\SystemBanner\Office`, registers it for COM with RegAsm, and loads it in Word, Excel and PowerPoint for every user (`HKLM\SOFTWARE\Microsoft\Office\<App>\Addins\SystemBannerOffice.Connect`). Restart Office.

To remove it, close Office and run `Install\RemoveSystemBannerOffice.bat` as an administrator.

--------
SETTINGS
--------

The add-in reads the SystemBanner key, `HKEY_LOCAL_MACHINE\SOFTWARE\Policies\SystemBanner`:

| Value | Type | Meaning |
|---|---|---|
| `Simple` | DWORD | The classification of the computer, same codes as SystemBanner: 1 to 6 for the United States, 11 to 52 for the other systems. It is the ceiling of the documents written on it. Not set or 0: no ceiling, IC markings. |
| `OfficeRequireClassification` | DWORD | 1: a document with no classification cannot be saved or printed. |
| `OfficeSpellOut` | DWORD | 1: banner lines spell control markings out (`TALENT KEYHOLE`) instead of using their abbreviations (`TK`). |
| `OfficeAnalyticProducts` | DWORD | 1: documents are disseminated analytic products, every classified portion must carry an FD&R marking (ICD 710 E.5). |
| `OfficeToolbar` | DWORD | 1: the SystemBanner toolbar in Office 2007 and later too (Add-ins tab), besides the ribbon. A user may also set it in `HKEY_CURRENT_USER\Software\SystemBanner`. |

---
USE
---

The **Classification** tab (Office 2007 and later) or the **SystemBanner** toolbar (Office 2000 to 2003) has four buttons:

- **Classify**: picks the classification of a document with no portion marks, or of an Excel workbook. The dialog shows the banner line and the portion mark as they are typed, and the reasons a combination is refused.
- **Portion mark**: marks the paragraph at the cursor, then rolls up the banner line.
- **Roll up**: rebuilds the banner line from every portion mark of the document, and reports unmarked paragraphs and the choices it made.
- **Authority block**: writes the classification authority block at the cursor in Word, and stores it in the document properties.

------------------------
WHAT THE ADD-IN CHECKS
------------------------

- Syntax: uppercase, category order, Register order inside a category, no `////` placeholders, `USA` first in REL TO, `SAR-` written once.
- Levels each marking allows (for example SI-GAMMA TOP SECRET only, FOUO, LIMDIS and SBU unclassified only).
- What each marking requires (HCS-O requires ORCON and NOFORN, RESERVE a 3 character compartment, GAMMA sub-compartments 4 letters, TK compartments NOFORN, EXDIS and NODIS NOFORN, JOINT REL TO its co-owners...).
- Markings that cannot be combined (NOFORN with REL TO, RELIDO, EYES ONLY or DISPLAY ONLY; ORCON with RELIDO; EXDIS with NODIS...).
- SIGMA numbers (14, 15, 18 and 20).

------
LIMITS
------

- **The Register is from 2016.** CUI (Controlled Unclassified Information) has since replaced FOUO and SBU in most uses; the add-in does not implement CUI markings yet.
- **Appendices A, B and C** of the Manual (non-US, NATO and UN values) and the Register annexes of trigraph and tetragraph codes are not published with it: country codes are checked for their form (3 or 4 letters), not against a list, and only the NAC activity of the Manual's own example (ISAF, `NIS`) has a portion mark.
- Where the Manual leaves room or contradicts itself, the roll-up reports its choice in its message: unclassified caveated portions in a classified document are handled as RELIDO (the Manual's `(U//FISA)` example), JOINT co-owners are added to the FGI list of a US banner (the rule, where the Manual's example leaves GBR out), foreign RESTRICTED rolls up to a US banner as CONFIDENTIAL.
- Portion marks are read at the start of paragraphs of the Word body and of PowerPoint text boxes; footnotes, text boxes in Word, PowerPoint notes and table cells in PowerPoint are not read. Excel workbooks are marked as a whole.
- Loaded by Word, Excel and PowerPoint 2016 (64 bit), ribbon and toolbar included, and driven through automation (see below). Office 2000 to 2010 still have to be tested on their own versions, see `TESTING.md`.

-------------------
BUILDING AND TESTING
-------------------

`build.bat` builds everything with the C# 3.0 compiler of the .NET Framework 3.5, against the .NET Framework 2.0 assemblies only, into `build` (or the folder given as argument; the old compiler fails on paths longer than about 200 characters):

- `SystemBannerOffice.Markings.dll`: the marking engine (model, parser, formatter, roll-up, validator, authority block), with no Office dependency.
- `SystemBannerOffice.dll`: the add-in, every Office call late bound, so it needs no primary interop assembly.
- `SystemBannerOffice.Tests.exe`: the tests of the engine, built on the examples of the Manual (section J and the notional pages of the marking templates).
- `OfficeCheck.exe [word] [excel] [powerpoint]`: drives the installed Office through automation with the add-in's document code, without installing the add-in: banners, portion marks, roll-up, authority block, document properties, save events.
- `OfficeCheck.exe load [word|excel|powerpoint]`: Office itself loads the installed add-in, as the COM add-ins dialog does, and shows the end of its log.
- `OfficeCheck.exe diag`: what Windows and Office see of the installation, written to `SystemBannerOffice_diag.txt` on the Desktop.

The add-in logs each start and every error to `%TEMP%\SystemBannerOffice.log`.

-------
CREDITS
-------

SystemBanner was created by Aaron Wawrzyniak, the Windows 2000 to 7 and Mac fork is by Cole. Stock classification colors are based on the GSA form color information collected by Frank Caviggia. Released under the MIT License.
