Roadmap
=======

Plan for the next working session, 25 October 2026. Order: the two bugs first (they block the tests on the VMs), then the German interface, then EU markings.

State on 24 October 2026: IC, French, German and British markings, English and French interface, 61 unit tests, every OfficeCheck check passed on Office 2016 (64 bit). Tested by hand on Office 2016. On the VMs: PowerPoint 2000 shows the toolbar; the two bugs below remain.

----------------------------------------------
1. Bug: no toolbar in Word 2000 and Excel 2000
----------------------------------------------

Seen on the Windows 2000 VM: the SystemBanner toolbar shows in PowerPoint 2000, not in Word 2000 nor Excel 2000.

**First, collect** (Windows 2000 VM, latest build):
1. Start Word 2000, then Excel 2000, close them normally.
2. Run `build\OfficeCheck.exe diag`, keep `SystemBannerOffice_diag.txt`.
3. Look at `%TEMP%\SystemBannerOffice.log` for `[WINWORD]` and `[EXCEL]` lines.

**Hypotheses, most likely first:**

| # | What the log shows | Hypothesis | Fix |
|---|---|---|---|
| A | `ERROR in CreateToolbar, Click event of Classify` (or another step) | One step fails, for example linking the Click event of a button in Word or Excel 2000. Everything sits in one `try`, and `Visible` is set last, so the toolbar is never shown. PowerPoint does not hit the failing step. | Show the toolbar before linking events; catch errors per button, so a failing event only disables one button; log each step. |
| B | `toolbar created with 5 buttons`, but nothing on screen | Word and Excel restore their own toolbar layout after the add-in starts (Normal.dot, Excel.xlb), and hide or move it. | Word: set `Application.CustomizationContext` to the Normal template before creating it. Both: set `Visible` and `Position` again on the first window activation, or create the toolbar on the first `WindowActivate`/`NewWorkbook` event. |
| C | No `[WINWORD]` or `[EXCEL]` line at all | Word and Excel 2000 never start the add-in: `LoadBehavior` turned to 2 after an old crash, or they do not read the `HKLM` Addins key. | Check `LoadBehavior` in the diag. Try `Install\SystemBannerOffice_currentuser.reg`. If it works, the installer also writes the `HKCU` key on Office 2000. |

**Done when:** the toolbar shows in Word, Excel and PowerPoint 2000, each of its 5 buttons opens its dialog, and the log has no error.

------------------------------------------------------
2. Bug: nothing happens on Classify in Office 2007
------------------------------------------------------

Seen on the Windows XP VM with Office 2007: the Classification tab shows, Classify does nothing.

**First, collect** (Windows XP VM):
1. In Word 2007: Office button > Word Options > Advanced > General, tick **Show add-in user interface errors**. Restart Word, click Classify, note any message.
2. After the click: is Word still responsive, or frozen (a dialog hidden behind its window)?
3. `%TEMP%\SystemBannerOffice.log` and `build\OfficeCheck.exe diag`.

**To add before testing:** a log line at the start of every command (`Run`), so the log tells whether the click reached the add-in.

**Hypotheses, most likely first:**

| # | Symptom | Hypothesis | Fix |
|---|---|---|---|
| A | Word frozen after the click | The dialog opens behind Word or off screen: its owner is `GetForegroundWindow()`, which may be the ribbon's own window in Office 2007. | Use the Office main window as the owner: `FindWindow("OpusApp")` for Word, `"XLMAIN"` for Excel, `"PPTFrameClass"` for PowerPoint, with the caption of the active window; bring the dialog to the front. |
| B | No log line for the click, a ribbon error message | Office 2007 does not find or call `OnAction` (callback signature). | Declare the `IRibbonControl` interface (`000C0395-0000-0000-C000-000000000046`) and take it as the callback parameter; give each button its own callback if needed. |
| C | `ERROR in Classify` in the log | An exception in the command, shown in a message box that is itself hidden (see A). | Fix the error the log names. |

**Done when:** on Office 2007, every button opens its dialog in front of Office, in Word, Excel and PowerPoint.

**Also test on the way:** Office 2003 and Office 2010 (32 and 64 bit), see `TESTING.md`.

------------------------
3. German interface
------------------------

Goal: dialogs, buttons and messages in German as well, for German users.

1. `Text.T(english, french)` becomes `Text.T(english, french, german)`; the 2 argument form stays during the change and falls back to English.
2. Language: `OfficeLanguage` = `de`, or an Office interface in German (LCID & 0x3FF = 0x07). The German markings (GEHEIM, VS-NfD...) stay as they are; the IC, French and British markings are never translated.
3. Translate file by file (about 300 messages: engine, then dialogs, then commands); these can be split between several people or agents, one file each.
4. **Review by a German speaker**, especially the security terms (Einstufung, Geheimhaltungsgrad, Herabstufung...).
5. Tests: the language test checks a German Office (0x0407) and a message in the three languages.

**Done when:** every message has its German text, reviewed, and the add-in follows a German Office.

------------------------------------
4. EU classified information (EUCI)
------------------------------------

Goal: mark documents with the classifications of the European Union.

**To read first** (primary sources, not summaries):
- Council Decision 2013/488/EU on the security rules for protecting EU classified information, and its annexes on marking.
- Commission Decision (EU, Euratom) 2015/444 on the security rules for protecting EU classified information.
- If needed, the rules for non classified sensitive information of the Commission (marking `SENSITIVE`, `LIMITED`), to decide whether they are in scope.

**Levels to check in those texts:** TRÈS SECRET UE/EU TOP SECRET, SECRET UE/EU SECRET, CONFIDENTIEL UE/EU CONFIDENTIAL, RESTREINT UE/EU RESTRICTED, and how they are written (bilingual form, abbreviations, place on the page, colors if any, releasability and other markings, portion marking, declassification).

**Steps:**
1. Write down the rules with their article numbers, like the other systems in the README; list what the texts leave open.
2. Engine `Eu.cs`: levels, markings, roll-up if the texts provide portion marking, validation, tests.
3. Stamps in Word, Excel and PowerPoint, dialog (in the three languages), system chooser, document properties.
4. Ceiling: SystemBanner has no EU code yet. Either EU markings are offered only with no `Simple` value, or new codes are added to SystemBanner itself (for example 61 to 64), in the SystemBanner_XP repository, with its Group Policy templates. To decide.
5. README section and TESTING.md checks.

**Done when:** an EU document can be classified, stamped and checked in Word, Excel and PowerPoint, with the rules documented and their sources.

--------------------------------
5. Outlook add-in (to evaluate)
--------------------------------

Not decided: a short study first, then go or no go with an estimate.

- **What it would do:** classify an email like a document (subject prefix, marking at the top and bottom of the body, message headers for mail gateways), with the same systems and the same ceiling, and block sending an email less classified than its attachments.
- **Risks to weigh:** Outlook 2000 to 2003 show security prompts when an add-in reads recipients or sends (Object Model Guard); Outlook 2000 has no ribbon and its message windows need their own toolbars; HTML, RTF and plain text bodies are marked differently.
- **Done when:** the study says what works on Outlook 2000 to 2010, what it costs, and the decision is taken.

----------------
6. Screenshots
----------------

Like SystemBanner (`Docs/Images`), the README shows what the add-in looks like.

- **Dialogs, generated:** a new `OfficeCheck.exe screenshots` mode opens each dialog filled with an example (system choice, IC, French, German, British, first page, authority block, downgrade) and saves it as a PNG with `Control.DrawToBitmap`, in English and French (and German after step 3), so they follow each build.
- **Office, captured by hand:** one marked page of each system (IC banner, SECRET SPÉCIAL FRANCE, GEHEIM, OFFICIAL-SENSITIVE), the ribbon tab in Office 2016 and 2007, the toolbar in Office 2000, Excel and PowerPoint; Office 2000 and 2007 once bugs 1 and 2 are fixed.
- **Where:** `Docs/Images`, named `<Office>_<system>_<what>.png` (for example `Word2016_France_Secret.png`), example content only, shown in the README under each system.
- **Done when:** the README shows the dialogs and one marked page of each system.

------------------
Still open, later
------------------

- CUI markings (DoDM 5200.48) for United States unclassified information, replacing FOUO.
- NATO and EU information inside French documents (the IGI equivalence table).
- A strong name for the two DLLs, to remove the RegAsm RA0000 warning.
