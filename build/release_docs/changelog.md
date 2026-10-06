<!--
Editor's notes for each release, newest first. Write the notes for work in
progress under "## Unreleased"; the build moves that section under the
version it ships in and starts a fresh empty "Unreleased". Use the
sub-headings Added / Corrected / Expanded / Restructured / Also, bullets
only, release-notes tone. The build appends its own measured comparison
with the previous release after these notes.
-->

## Unreleased

## v20261006c -- 6 October 2026

### Added
- Major U.S. Christian Denominations: a third table, "Other divisive positions," in every one of the 27 entries, covering 13 further questions that have split churches (Scripture and Tradition; the Trinity and the person of Christ; predestination and free will; sanctification and holiness; hell and the fate of the lost; Mary, the saints and prayer for the dead; the day, style and liturgy of worship; church and state; race and the church; creation and the age of the earth; tithing and the prosperity gospel; Bible translations and KJV-onlyism; and marriage roles), each with a source label -- 351 new position cells, bringing the tracked positions per body to 31.
- Prominent English Study Bibles: the same "Other divisive positions" table in every one of the 27 entries, recording what each volume's notes say or assume on the 13 questions (most visibly at Genesis 1, Romans 9, Ephesians 5, Matthew 25 and Revelation 20) and, for the translations row, the volume's own base text -- a further 351 cells.
- Remaining-verifications workbook: both matrix sheets gain the 13 new columns, and the ledger gains rows 97-109 (denominations) and 110-122 (study Bibles), one per new category, filled by the build from the live tables, plus row 123 for the dates and attributions in the new section.

- New section, *Major Dividing Issues Among Believers* (130): one chapter for each of the 31 tracked positions -- what divides, the main answers, the passages each side rests on (pointing into the two catalogs by entry ID), where every denomination and study Bible in this book falls, the dated splits, and how to study it. The where-they-fall tables are generated at build from the position tables, so they cannot drift from the entries. References for Further Reading moves to 140.

### Corrected
- Remaining-verifications workbook: the amber "not independently documented" flag now also catches the study-Bible label "General pattern of the editorial tradition, not independently documented in the notes", which the earlier pattern missed, so those cells are no longer shown as if they had a cited source.

### Also
- Section introductions for the Denominations and Study Bibles sections now describe all three tables and all 31 positions.

## v20261003a -- 3 October 2026

### Added
- Front cover illustration page.
- Scripture Index: every verse cited in the two catalogs, generated from the entries themselves so it cannot drift from them.
- Entry IDs on every Reportedly Contradicting Passages entry (e.g. MATT-001), giving each a stable citation handle.
- New Introduction chapter *How to Study a Passage*: context at three distances, genre, comparing translations, weighing a variant without Greek, common traps, and a short procedure.
- *Character of Each Source Manuscript Tradition* gains a worked example -- weighing the evidence for the ending of Mark (16:9-20).
- Two denominations added (African Methodist Episcopal Zion Church; United Pentecostal Church International), bringing the section to 27, now ranked by 2020 U.S. Religion Census adherents.
- Study Bibles: a doctrinal-positions table for all 27 volumes, matching the denominations section; ethical-and-lifestyle commentary reorganised as a positions table with sources.
- References for Further Reading now opens with a statement of method: how translation wording, manuscript evidence, harmonizations and denominational/study-Bible positions were checked, and that drafting was AI-assisted with single-editor review.

### Corrected
- All 78 "not independently confirmed" notes resolved: every one of the 27 translations has been checked against its own text at every catalogued variant.
- Ruth 3:15: the NIV and NLT read "he went" (Boaz), not "she"; the "he" reading now leads 13 translations to 14.
- 1 Samuel 10:27: the NRSV-CE, the ERV (unbracketed) and the NLT (in brackets) carry the Qumran paragraph about Nahash the Ammonite; earlier text named only the NRSV-CE.
- Ephesians 1:1: the RSV omits "in Ephesus" from its main text; NABRE and the NET Bible print it in brackets.
- Psalm 22:16: the CPDV reads "pierced," not Douay-Rheims' "dug."
- Leviticus 25/26 and 2 Chronicles 36:21: the ERV paraphrases "sabbath" to "rest" in Leviticus while keeping it in Chronicles; recorded.
- Christian and Missionary Alliance and Presbyterian Church in America: exact 2020 Census counts replace estimates (414,360 and 372,696 adherents).
- World English Bible: online editions differ at Leviticus 17:4 and 1 Samuel 10:27; noted in both entries.
- Esther: confirmed that none of the 27 tracked translations introduces a divine name into chapters 1-10.
- Deuterocanon: CPDV confirmed to follow the Vulgate/Douay-Rheims form in Tobit, Judith, Wisdom and Sirach 33:1.
- The Passover-lamb and law-of-the-king entries (Deuteronomy) and the "Who killed Goliath" entry (1 Samuel) consolidated with their Exodus, 1 Kings and 2 Samuel treatments.

### Expanded
- Manuscript and Translation Differences: every side note now names all 27 translations by reading (TRB's bracketed renderings recorded throughout).
- Reportedly Contradicting Passages: chapter introductions added or extended; translation comparisons enlarged in Genesis, Exodus, 1 Kings, Matthew, John, Psalms and some 30 other chapters.
- Top Christian Denominations re-ordered and re-checked; sources named for every ethical-and-lifestyle position.
- *Purpose and Scope* and *A Note on Method and Verification* rewritten for the completed verification.

### Also
- A Remaining Verifications spreadsheet now accompanies every release.
- Release files are renamed *a_Bible_study_primer_vYYYYMMDDx*; earlier releases keep their old names.

## v20260919f -- 19 September 2026

### Added
- Two new Introduction chapters: *What Is Meant by the Word of God* and *What Is Meant by an Inerrant Word of God*.
- Title page note to readers on what to report via GitHub Issues.
- Top Christian Denominations (new section): the 25 largest U.S. Christian bodies with doctrinal and ethical-and-lifestyle tables, each position labelled with its source.
- Top Study Bibles (new section): 27 prominent study Bibles with base translation, editorial credit, denominational leaning and ethical-and-lifestyle commentary.
- Eight new translation histories: The Readable Bible, the Complete Jewish Bible, the King James tradition, the ASV and its descendants, the RSV-ESV tradition, the Catholic RSV editions, the Vulgate tradition in English, and the 19th-century literalist translations.

### Expanded
- Tracked translations: 10 to 27 (AKJV, ASV, BSB, CJB, CPDV, Darby, Douay-Rheims, ERV, NABRE, NET Bible, NRSV-CE, RSV, Smith's Literal Translation, TRB, WEB, Webster's Bible and YLT added).
- Popular Bible Translations now ranks all 27.
- Manuscript and Translation Differences: 142 to 202 catalogued variants, readings grouped into "camps" naming every translation.
- Reportedly Contradicting Passages: "How the translations render it" rewritten for all 27 in every entry.
- References for Further Reading extended.

### Corrected
- Zechariah 11:13 and 12:10 side notes rewritten with each translation's actual reading.
- Where a translation's reading could not be confirmed directly, the entry now says so rather than implying it was checked.

### Restructured
- Translation histories consolidated from 26 single-translation pages into 16 family histories.
- Sections renumbered: Manuscript and Translation Differences 090, Reportedly Contradicting Passages 100, Denominations 110, Study Bibles 120, References 130.
