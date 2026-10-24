# EU classified information (EUCI): marking rules from the primary texts

Research date: 24 October 2026. All quotes and references below were checked against the Official Journal texts served by the EU Publications Office (Cellar, the repository behind EUR-Lex), English and French language versions. EUR-Lex itself blocked automated access, so the texts were fetched from `publications.europa.eu/resource/celex/<CELEX>` (same OJ content).

## Sources

| Short name | Act | CELEX | Status found |
|---|---|---|---|
| **CSR** (Council security rules) | Council Decision 2013/488/EU of 23 September 2013 on the security rules for protecting EU classified information, OJ L 274, 15.10.2013, p. 1 | 32013D0488 | In force. Amended four times, **only Appendices B (equivalence table) and C (list of NSAs)**: Decisions 2014/233/EU, (EU) 2019/2247, (EU) 2021/1075, (EU) 2025/1540. Articles and Annex III (marking) are unchanged since 2013. |
| **CSD** (Commission security decision) | Commission Decision (EU, Euratom) 2015/444 of 13 March 2015 on the security rules for protecting EU classified information, OJ L 72, 17.3.2015, p. 53 | 32015D0444 | In force. No amending or repealing act recorded in Cellar. |
| 2015/443 | Commission Decision (EU, Euratom) 2015/443 of 13 March 2015 on Security in the Commission, OJ L 72, 17.3.2015, p. 41 | 32015D0443 | Used only for "sensitive non-classified information". |
| 2023/2841 | Regulation (EU, Euratom) 2023/2841 (cybersecurity in Union entities) | 32023R2841 | Does **not** apply to systems handling EUCI (Art. 2(3)), except Art. 13(8). No marking rules. |
| Info-security proposal | Proposal for a Regulation on information security in the Union institutions, COM(2022) 119 | 52022PC0119 | **Not found as an adopted act** in a Cellar search of 2020-2026 acts (titles with "information security" / "classified information"). Council documents from July 2025 still show a revised draft. **To verify** before release: if adopted, it may replace the CSD rules. |

Structure note: in the CSR the marking rules sit in **Article 2, Article 3, Article 9 and Annex III** ("Management of classified information"), with electronic rules in **Annex IV**. In the CSD the same rules are in the body: **Articles 3, 4, 21 to 30, 38, 57**. The wording of the two texts is almost identical. Below, "CSR Annex III pt 13" means point (paragraph) 13 of Annex III.

---

## 1. The four levels and their exact written form

| Level | Written form (identical in EN and FR texts) | Paragraph abbreviation |
|---|---|---|
| Top secret | `TRÈS SECRET UE/EU TOP SECRET` | `TS-UE/EU-TS` |
| Secret | `SECRET UE/EU SECRET` | `S-UE/EU-S` |
| Confidential | `CONFIDENTIEL UE/EU CONFIDENTIAL` | `C-UE/EU-C` |
| Restricted | `RESTREINT UE/EU RESTRICTED` | `R-UE/EU-R` |

Sources: CSR Art. 2(2)(a)-(d); CSD Art. 3(2)(a)-(d). Abbreviations: CSR Annex III pt 12; CSD Art. 24(2).

Rules and observations:

- **EUCI must bear one of these markings.** "EUCI shall bear a security classification marking in accordance with paragraph 2" (CSR Art. 2(3); CSD Art. 3(3)).
- **The bilingual string is the marking itself, not a translation.** The English, French and German OJ versions all print exactly the same string (`TRÈS SECRET UE/EU TOP SECRET`, etc.). The French version does not use a French-only form, and the English version does not use an English-only form. Checked: CSR Art. 2(2) EN, FR, DE; CSD Art. 3(2) EN, FR.
- **Order: French part first, then English part**, as printed in every version: `<FR level> UE` + `/` + `EU <EN level>`.
- **Separator:** a single slash with no spaces (`UE/EU`). Inside each half, a normal space.
- **Capitals:** all capitals in every occurrence. The texts contain no explicit rule on capitals; the capitals come from the printed form only.
- **Accent:** the CSR prints `TRÈS` (with grave accent) in EN, FR and DE. The CSD English version prints `TRES` without accent throughout (14 occurrences, e.g. Art. 3(2)(a), Art. 24(2)); the CSD French version prints `TRÈS`. One unaccented `TRES` also appears in the CSR French version, Annex III pt 24. These look like typographic variants, not a separate form. **Recommendation for the add-in: use `TRÈS`** (the Council text, the French text of both decisions, and the Appendix B equivalence table of Decision (EU) 2025/1540 all use it).
- **No other wording is admitted** for the four levels: the texts give a closed list (CSR Art. 2(2): "shall be classified at one of the following levels").
- Unclassified is not a level. The texts define no "UNCLASSIFIED" or "NON CLASSIFIÉ" marking.

Equivalence with the national markings already supported (CSR Appendix B, as replaced by Decision (EU) 2025/1540, Annex I):

| EU | France | Germany |
|---|---|---|
| TRÈS SECRET UE/EU TOP SECRET | TRÈS SECRET (and TRÈS SECRET DÉFENSE before 1 July 2021) | STRENG GEHEIM |
| SECRET UE/EU SECRET | SECRET (and SECRET DÉFENSE before 1 July 2021) | GEHEIM |
| CONFIDENTIEL UE/EU CONFIDENTIAL | CONFIDENTIEL DÉFENSE (pre-2021 only; France protects it as SECRET, note 2b) | VS-VERTRAULICH (printed "VS ( 1 ) VERTRAULICH", VS = Verschlusssache) |
| RESTREINT UE/EU RESTRICTED | No French equivalent (note 2c: France does not use RESTREINT nationally) | VS – NUR FÜR DEN DIENSTGEBRAUCH |

The United Kingdom was removed from Appendix B by Decision (EU) 2021/1075 (recital 4): there is no EU-UK equivalence in this table any more (a separate EU-UK security of information agreement exists, Decision (EU) 2021/689; not studied here).

## 2. Where the marking goes on the page

- **Every page must carry the classification level.** "each page shall be marked clearly with the classification level" (CSR Annex III pt 13(a); CSD Art. 25(1)(a)).
- **The marking must be clear and correct whatever the medium**: paper, oral, electronic or other (CSR Annex III pt 5; CSD Art. 22(4)).
- **Top/bottom, centring, colour, font size: not specified.** Neither text says where on the page the marking goes, nor any colour, size, font or bold requirement. Searched both texts in EN and FR: no rule found.
- Where pt 13 cannot be applied (e.g. objects, media), "other appropriate measures" apply under security guidelines (CSR Annex III pt 14, guidelines under Art. 6(2)) or Commission implementing rules (CSD Art. 25(2)). Those guidelines and implementing rules are **not published in the OJ** and were not available as primary sources. **Not verified.**

## 3. Additional markings

### What the texts allow

- The general rule: EUCI "may bear additional markings to designate the field of activity to which it relates, identify the originator, limit distribution, restrict use or indicate releasability" (CSR Art. 2(3)). The CSD adds that these "are not classification markings" (CSD Art. 3(3)).
- The list of examples (CSR Annex III pt 10; CSD Art. 23):
  - (a) an identifier to designate the **originator**;
  - (b) **caveats, code-words or acronyms** for the field of activity, a need-to-know distribution, or restrictions on use;
  - (c) **releasability markings**;
  - (d) where applicable, the **date or specific event** after which it may be downgraded or declassified.
- French terms (CSR FR Annex III pt 10; CSD FR Art. 23): "identifiant désignant l'autorité d'origine", "marquages restrictifs, des mots-codes ou des acronymes", "marquages relatifs à la communicabilité", "la date ou l'événement particulier à partir desquels elles peuvent être déclassées ou déclassifiées".
- These markings are **optional** ("may bear", "such as"): the list is open.
- **Only the originator can change or remove them**: no downgrading, declassification, or change/removal of the Art. 2(3) markings "without the prior written consent of the originator" (CSR Art. 3(2)). The CSD version covers only the classification markings of Art. 3(2), and is "without prejudice to Article 26" (CSD Art. 4(2)).

### Releasability

- The texts require a releasability marking **when EUCI is released to a third State or international organisation**: the document "shall bear a releasability marking indicating the third State or international organisation to which it has been released" (CSR Annex VI pt 36; CSD Art. 57(3)).
- **No format is given.** The texts contain neither "REL TO" nor "RELEASABLE TO" nor any country-code convention. **Left open.**

### "EU/EURATOM" designations

- Not a marking in these texts. The only Euratom mention is "Euratom classified information" as a separate category in a transitional rule (CSD Art. 59(2)) and in 2015/443 Art. 9. No "EURATOM" marking format is defined. **Not in scope / not found.**

### Covering letter detached from enclosures

- A covering letter or note is classified at least as high as its highest enclosure. The originator must show its level when detached, e.g. (CSR Annex III pt 9; CSD Art. 22(8)):
  - EN: `CONFIDENTIEL UE/EU CONFIDENTIAL` then `Without attachment(s) RESTREINT UE/EU RESTRICTED`
  - FR: `CONFIDENTIEL UE/EU CONFIDENTIAL` then `Sans pièce(s) jointe(s) RESTREINT UE/EU RESTRICTED`
- This is given as an example ("e.g." / "par exemple"), so the exact wording is not mandatory.

### Non-classified markings ("LIMITE", "SENSITIVE")

- **Out of scope of the CSR and CSD.** Neither text contains "LIMITE" or a "SENSITIVE" marking (searched EN and FR).
- The Commission's **"sensitive non-classified information"** is defined in Decision 2015/443, Art. 9(5)(b). Art. 9(6) says it is identified "by a security marking and corresponding handling instructions" approved by the Director-General for Human Resources and Security, "when deemed necessary". **The marking itself (e.g. "SENSITIVE") is not in the OJ.** Not verified.
- The Council "LIMITE" distribution marking is not in Decision 2013/488/EU. Its source (Council internal rules on handling non-classified documents) was not checked. **Not verified.**

## 4. Portion (paragraph) marking

- **Optional, not required.** "Standardised abbreviated classification markings may be used to indicate the classification level of individual paragraphs of a text" (CSR Annex III pt 11; CSD Art. 24(1)).
- **Abbreviations never replace the full marking**: "Abbreviations shall not replace the full classification markings" (same references). The page marking in full stays mandatory.
- **Scope:** "sections or blocks of text of less than a single page" (CSR Annex III pt 12; CSD Art. 24(2)).
- **Form (closed list, identical in EN and FR):** `TS-UE/EU-TS`, `S-UE/EU-S`, `C-UE/EU-C`, `R-UE/EU-R`.
- **Brackets are not specified.** The texts print the abbreviations without parentheses. Whether to write `(R-UE/EU-R)` or `R-UE/EU-R`, and where (start of paragraph, end), is **left open**.
- Related mandatory principle: individual parts of a document (pages, paragraphs, sections, annexes, appendices, attachments, enclosures) "may require different classifications and shall be marked accordingly, including when stored in electronic form" (CSR Annex III pt 6). The CSD says "be marked accordingly" (CSD Art. 22(5)), a softer wording. So if parts have different levels, marking them is required by the CSR, but using the abbreviations to do it is optional.
- **Overall level** of a document or file is at least as high as its most highly classified part; aggregation may justify a higher level (CSR Annex III pt 7; CSD Art. 22(6); also CSR Art. 4(4)).
- Documents with mixed levels should, "to the extent possible", be structured so parts can be identified and detached (CSR Annex III pt 8; CSD Art. 22(7)).
- No portion marking is defined for unclassified paragraphs (no "U" or "NC" abbreviation exists in the texts).

## 5. Downgrading and declassification

- **Who decides:** the originator. No downgrading or declassification without the originator's prior written consent (CSR Art. 3(2); CSD Art. 4(2)). The originator also sets the level initially (CSR Annex III pt 3; CSD Art. 22(2)).
- **What to mark:** at creation, the originator "shall indicate, where possible" whether the EUCI can be downgraded or declassified **on a given date or following a specific event** (CSR Annex III pt 15; CSD Art. 26(1)). The CSR adds "in particular for information classified RESTREINT UE/EU RESTRICTED". This is an additional marking under CSR Annex III pt 10(d) / CSD Art. 23(d).
- It is "where possible", so **not strictly mandatory**.
- **Review:** the Council secretariat (GSC) reviews its EUCI at least every five years; no review is needed if the originator marked automatic downgrading/declassification from the outset (CSR Annex III pt 16; CSD Art. 26(2) for Commission departments).
- **Commission-only automatic rule:** RESTREINT UE/EU RESTRICTED information originated by the Commission is considered automatically declassified after **thirty years** (CSD Art. 26(3), by reference to Regulation 354/83 as amended by Regulation 1700/2003).
- Definitions: "Downgrading" is a reduction in the level; "Declassification" is the removal of any classification (CSR Appendix A; CSD Art. 1). French: "déclassement" and "déclassification".
- **No wording or format** for the downgrading instruction is given (e.g. "Declassify on: ..."). **Left open.**

## 6. Registration, copy numbers, page numbers, annexes, cover pages

Required for **every** EU classified document at creation (CSR Annex III pt 13; CSD Art. 25(1)):

| Rule | Level | Reference |
|---|---|---|
| Classification on each page | All levels | CSR Annex III pt 13(a); CSD Art. 25(1)(a) |
| Each page numbered | All levels | CSR Annex III pt 13(b); CSD Art. 25(1)(b) |
| A **reference number** (CSR) / **registration number** (CSD) and a subject; the subject is not itself classified unless marked as such | All levels | CSR Annex III pt 13(c); CSD Art. 25(1)(c) |
| Document dated | All levels | CSR Annex III pt 13(d); CSD Art. 25(1)(d) |
| **Copy number on every page**, if distributed in several copies | SECRET UE/EU SECRET and above | CSR Annex III pt 13(e); CSD Art. 25(1)(e) |

Registration:

- CONFIDENTIEL UE/EU CONFIDENTIAL and above must be **registered for security purposes** before distribution and on receipt (CSR Art. 9(2); CSD Art. 21(2)); registered in designated registries when it arrives at or leaves an entity (CSR Annex III pt 19; CSD Art. 29(2)).
- TRÈS SECRET UE/EU TOP SECRET is registered in **designated** registries (CSR Art. 9(2)); dedicated central registry (CSR Annex III pt 23-24; CSD Art. 27(3)-(4)).
- In a CIS (IT system), registration may be done by processes within the system itself (CSR Annex III pt 21; CSD Art. 29(3)).
- RESTREINT UE/EU RESTRICTED: no security registration required by these texts.

Copying and translation:

- TRÈS SECRET UE/EU TOP SECRET: no copy or translation without the originator's prior written consent (CSR Annex III pt 25; CSD Art. 30(1)).
- SECRET and below: may be copied or translated on instruction from the holder, unless the originator imposed caveats (CSR Annex III pt 26; CSD Art. 30(2)).
- Copies and translations get the same protection as the original (CSR Annex III pt 27; CSD Art. 30(3)).

Annexes and parts:

- Annexes, appendices, attachments and enclosures may have their own level and must then be marked accordingly (CSR Annex III pt 6; CSD Art. 22(5)).
- Covering letter rule: see section 3.

Cover pages:

- **No cover page (cover sheet) requirement** exists in either text. For carriage within a building, EUCI must be "covered in order to prevent observation of its contents" (CSR Annex III pt 30), which concerns physical carriage, not a document cover page. **Left open.**

## 7. Electronic documents

- The classification must be clearly and correctly indicated in electronic form too (CSR Annex III pt 5; CSD Art. 22(4)).
- Parts with different levels must be marked "including when stored in electronic form" (CSR Annex III pt 6; CSD Art. 22(5)).
- Registration may be done inside the CIS (CSR Annex III pt 21; CSD Art. 29(3)).
- **Emergency exception (important for an add-in):** when EUCI is sent in an emergency with lower-grade or no encryption, it "shall not bear any markings or indications distinguishing it from information which is unclassified"; recipients are told the level by other means (CSR Annex IV pt 41, referring to pt 39-40; CSD Art. 38(3)). The add-in should allow a user to produce an unmarked document deliberately, or at least not prevent it.
- **No rule on metadata, file properties, file names, e-mail subject lines or banners on screen.** Searched both texts. **Left open.**
- Regulation 2023/2841 (cybersecurity) does not apply to systems handling EUCI (Art. 2(3)), so it adds no marking rule.

---

## Left open by the texts (the add-in must decide)

1. **Position, size, colour and font of the page marking.** The texts only say "each page ... marked clearly". Top only, bottom only, or both, centred or not: not specified. Suggested default: centred header and footer, same as the national schemes already supported, configurable.
2. **Releasability format.** A marking naming the recipient third State or organisation is required on release (CSR Annex VI pt 36; CSD Art. 57(3)), but no syntax is given ("REL TO", "RELEASABLE TO", country codes, French form "COMMUNICABLE À"...).
3. **Portion marking syntax.** The abbreviations are fixed, but parentheses, position in the paragraph, and whether to mark unclassified paragraphs are not.
4. **Accent on TRÈS.** Both `TRÈS` and `TRES` appear in the OJ (CSD English version uses `TRES`). Pick one (recommended: `TRÈS`) and decide whether to accept the other on input.
5. **Downgrading / declassification instruction wording** (date or event): required "where possible", format free.
6. **Originator identifier, caveats, code-words**: allowed, open list, no format, no vocabulary.
7. **Copy number and reference/registration number format**, and whether they sit in the header or footer. Copy number is required on every page for SECRET and above when several copies are made.
8. **Page numbering format** ("page x of y" or not): only "each page shall be numbered".
9. **Cover pages**: not required by the texts; optional feature.
10. **Electronic metadata** (document properties, custom XML, file names, e-mail subjects): no rule.
11. **Non-classified markings** ("LIMITE" for the Council, "SENSITIVE" or similar for the Commission): outside these texts; their sources were not verified. If supported, they must not be presented as EUCI levels.
12. **Which institution's rules to follow.** The CSR binds the Council, its secretariat and Member States handling EUCI; the CSD binds the Commission. Marking rules are the same in substance, with small differences (reference vs registration number; CSD Art. 26(3) 30-year rule; CSR Annex III pt 6 "shall be marked" vs CSD Art. 22(5) "be marked").

## Not verified in a primary text

- The Council security policies and guidelines referred to in CSR Art. 3(3) and Art. 6(2) (practical classification guide, guidelines on marking where pt 13 cannot apply): not published in the OJ.
- The Commission implementing rules and security notices under CSD Art. 60: not published in the OJ.
- Whether the proposed Regulation on information security (COM(2022) 119) was adopted after the Cellar search date. None found as of 24 October 2026; check EUR-Lex before release.
- The source and form of the Council "LIMITE" marking and the Commission "SENSITIVE" marking.
- The German (DE) language version was checked only for the level strings and the `R-UE/EU-R` abbreviation, not paragraph by paragraph.
