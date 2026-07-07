# Requirements Document

## Introduction

This feature is a web-based tool that converts an Indian Customs **Bill of Entry (BOE)** PDF into a
specific, pre-defined Excel workbook (the "CTN" format) used by the importer's accounting team. Today
this conversion is fully manual: a staff member reads the BOE PDF, copies each line item and header
field into a spreadsheet, simplifies item names, and computes landed-cost figures (customs duty, IGST,
surcharge, land cost with/without GST) by hand. The tool automates that conversion.

The scope of this document is **Milestone 1** only: read a BOE PDF accurately and produce the target
Excel in the exact layout of the supplied sample (`1357 ctn llp.xlsx`). Parsing accuracy is the primary,
non-negotiable requirement — every field, including anomalies and outliers, must be captured faithfully,
and no line item may be silently dropped or altered.

**Out of scope (future Milestone B):** importing/exporting the resulting data into the Tally accounting
application (whether through Tally APIs or a Tally-importable Excel format). Tally integration is noted
here only to confirm it is deliberately excluded from Milestone 1.

The reference inputs/outputs used to derive these requirements are:
- Source BOE PDF: `205090022062026INNSA1BE0230620261842.pdf` (30 pages, 45 line items)
- Target Excel: `1357 ctn llp.xlsx` (Sheet1, header block + 45-row line-item table + totals + auxiliary sections)

## Glossary

- **BOE (Bill of Entry)**: A customs clearance document filed for imported goods. Source input for this tool.
- **Converter**: The overall system being built that transforms a BOE PDF into the target Excel workbook.
- **PDF_Parser**: The component that reads a BOE PDF and extracts structured header and line-item data.
- **Excel_Generator**: The component that writes extracted and computed data into the target Excel layout.
- **Value_Calculator**: The component that computes derived numeric fields (duties, IGST, land cost, etc.).
- **Web_Interface**: The browser-based UI through which a User uploads a BOE PDF and downloads the Excel.
- **User**: The importer's staff member who operates the Converter.
- **Line_Item**: One imported product entry in the BOE (identified by its Item Serial Number).
- **Header_Block**: The top section of the target Excel (rows 1-8) containing document-level fields
  (Company name, Party Name, USD Rate, Details, Invoice No/Date, BE No/Date, B/L No/Date, etc.).
- **Item_Table**: The line-item table in the target Excel beginning at the header row (row 12).
- **Totals_Row**: The summary row in the target Excel that aggregates numeric columns (row 61 in the sample).
- **Grand_Total**: A single aggregated column total written in the Totals_Row. The Grand_Totals referenced by
  the verification feature are: column P (TOTAL Custom Duty = per-line customs duty excluding IGST summed =
  `Totals.total_customs_duty`), column Q (GST = per-line IGST summed = `Totals.total_igst`), and column S
  (total custom duty = per-line combined duty summed = `Totals.total_customs_duty + Totals.total_igst`).
- **Duty_Summary**: The Part I "C. DUTY SUMMARY" section of the BOE that prints document-level declared duty
  totals, including BCD, ACD, SWS, NCCD, ADD, CVD, IGST, G.CESS, TOT.ASS VAL, TOTAL DUTY, and TOT.AMOUNT.
- **Declared_TOTAL_DUTY**: The TOTAL DUTY value printed in the BOE Duty_Summary (465302 in the sample); the
  sum of all customs duties plus IGST.
- **Declared_IGST**: The IGST value printed in the BOE Duty_Summary (258391 in the sample).
- **Declared_BCD**: The BCD (Basic Customs Duty) total value printed in the BOE Duty_Summary (184250.2 in the sample).
- **Declared_SWS**: The SWS (Social Welfare Surcharge) total value printed in the BOE Duty_Summary (18810.2 in the sample).
- **CTH / HSN Code**: Customs Tariff Heading / Harmonized System Nomenclature classification code for an item.
- **Assessable_Value**: The customs-assessed value (INR) of a Line_Item, shown as "ASSESS VALUE" in the BOE.
- **BCD (Basic Customs Duty)**: A customs duty levied as a percentage of the Assessable_Value.
- **SWS (Social Welfare Surcharge)**: A surcharge levied at 10% of the BCD amount.
- **IGST (Integrated GST)**: Goods and Services Tax charged on (Assessable_Value + customs duties).
- **USD_Rate**: The exchange rate (INR per USD) applied to convert invoice USD values to INR.
- **Unit_Price_USD**: The per-unit invoice price in USD (the BOE "UPI - Unit Price Invoiced" field).
- **Land_Cost**: The computed landed cost of an item, expressed both excluding and including GST.
- **Tally**: The external accounting application targeted by the out-of-scope future Milestone B.

## Requirements

### Requirement 1: Upload and convert a BOE PDF through a web interface

**User Story:** As a User, I want to upload a BOE PDF in a web browser and receive the converted Excel file, so that I no longer perform the conversion manually.

#### Acceptance Criteria

1. THE Web_Interface SHALL provide a control that accepts exactly one PDF file per submission as input.
2. IF the uploaded file exceeds 50 MB, THEN THE Converter SHALL reject the file, display a message stating that the file exceeds the 50 MB size limit, and SHALL NOT produce an output file.
3. WHEN a User submits a valid BOE PDF, where a valid BOE PDF is a readable PDF whose extracted content is recognized as a Bill of Entry containing the fields required to populate the sample CTN workbook layout, THE Converter SHALL produce a downloadable Excel workbook in the target CTN layout within 60 seconds of submission.
4. WHILE a conversion is in progress, THE Web_Interface SHALL display a processing indicator to the User, and SHALL continue displaying it until the conversion completes or an error is reported.
5. IF the uploaded file is not a readable PDF, THEN THE Converter SHALL reject the file, display a message stating that a valid PDF is required, and SHALL NOT produce an output file.
6. IF the uploaded PDF is not recognizable as a Bill of Entry, THEN THE Converter SHALL display a message stating that the document is not a recognized Bill of Entry and SHALL NOT produce an output file.
7. IF conversion fails after the uploaded PDF is recognized as a Bill of Entry, THEN THE Converter SHALL display a message indicating that the conversion could not be completed and SHALL NOT make any partial or incomplete Excel workbook available for download.
8. WHEN conversion completes, THE Web_Interface SHALL allow the User to download the generated Excel workbook.

### Requirement 2: Accurately extract all Bill of Entry header fields

**User Story:** As a User, I want every document-level field read correctly from the BOE, so that the Excel header block matches the source document.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract the Bill of Entry Number from the BOE.
2. THE PDF_Parser SHALL extract the Bill of Entry Date from the BOE.
3. THE PDF_Parser SHALL extract the Invoice Number from the BOE.
4. THE PDF_Parser SHALL extract the Invoice Date from the BOE.
5. THE PDF_Parser SHALL extract the total Invoice Amount and its currency from the BOE.
6. THE PDF_Parser SHALL extract the total number of packages (PKG / CTN count) from the BOE.
7. THE PDF_Parser SHALL extract the supplier/exporter name (Party Name) from the BOE.
8. THE PDF_Parser SHALL extract the container details (container number and count) from the BOE.
9. WHERE a Bill of Lading number and date are present in the BOE, THE PDF_Parser SHALL extract the Bill of Lading Number and Bill of Lading Date.
10. IF a header field required by criteria 1 through 8 cannot be located in the BOE, or its printed characters cannot be resolved into a value, THEN THE Converter SHALL record that field as missing, SHALL report each such missing field to the User identified by its field name, and SHALL NOT substitute an inferred or default value.
11. THE PDF_Parser SHALL record each extracted header field as the value printed in the BOE, preserving its characters, digits, and currency symbol without reformatting, truncating, rounding, or inferring content.

### Requirement 3: Accurately extract all Bill of Entry line items

**User Story:** As a User, I want every line item read correctly and completely, so that no product, quantity, or value is dropped or altered.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract one record for each Line_Item present in the BOE.
2. WHEN the BOE declares a total item count, THE PDF_Parser SHALL extract a number of Line_Item records equal to that declared count.
3. IF the number of extracted Line_Item records does not equal the BOE's declared item count, THEN THE Converter SHALL report the discrepancy to the User, stating both the BOE's declared item count and the number of records extracted, and SHALL NOT present the output as complete.
4. FOR each Line_Item, THE PDF_Parser SHALL extract the Item Serial Number.
5. FOR each Line_Item, THE PDF_Parser SHALL extract the CTH / HSN Code.
6. FOR each Line_Item, THE PDF_Parser SHALL extract the complete item description text verbatim, without truncation, abbreviation, or omission.
7. FOR each Line_Item, THE PDF_Parser SHALL extract the Unit_Price_USD.
8. FOR each Line_Item, THE PDF_Parser SHALL extract the declared quantity and the unit of quantity.
9. FOR each Line_Item, THE PDF_Parser SHALL extract the Assessable_Value.
10. FOR each Line_Item, THE PDF_Parser SHALL extract the BCD rate and BCD amount.
11. FOR each Line_Item, THE PDF_Parser SHALL extract the IGST rate.
12. FOR each Line_Item, THE PDF_Parser SHALL extract the total duty amount.
13. WHERE a Line_Item spans more than one PDF page, THE PDF_Parser SHALL merge the Line_Item's fields from all spanned pages into a single record with each field value appearing exactly once, without duplication or omission.
14. WHERE a Line_Item has an exemption notification that produces a zero duty rate, THE PDF_Parser SHALL extract the corresponding rate and amount as the numeric value zero rather than leaving them blank.
15. IF a Line_Item field required by criteria 4 through 12 cannot be located in the BOE, or its printed characters cannot be resolved into a value, THEN THE Converter SHALL record that field as missing, SHALL report it to the User identified by its Line_Item and field name, SHALL NOT substitute an inferred or default value, and SHALL NOT drop the Line_Item.

### Requirement 4: Populate the Excel header block from extracted data

**User Story:** As a User, I want the Excel header block filled from the BOE, so that the output identifies the shipment correctly.

#### Acceptance Criteria

1. THE Excel_Generator SHALL write the extracted supplier name into the Party Name field of the Header_Block.
2. THE Excel_Generator SHALL write the extracted Invoice Number and Invoice Date into the Header_Block exactly as extracted, without reformatting the number or date values.
3. THE Excel_Generator SHALL write the extracted Bill of Entry Number and Bill of Entry Date into the Header_Block exactly as extracted, without reformatting the number or date values.
4. WHERE a Bill of Lading Number and Bill of Lading Date are present in the extracted data, THE Excel_Generator SHALL write them into the Header_Block exactly as extracted, without reformatting the number or date values.
5. THE Excel_Generator SHALL write the USD_Rate into the Header_Block exactly as extracted, without altering its numeric precision.
6. THE Excel_Generator SHALL write the package/container details (e.g., "CO-32 CTN-1357") into the Details field of the Header_Block.
7. THE Excel_Generator SHALL write the importer company name into the Company name field of the Header_Block.
8. WHERE a Header_Block field has no corresponding value available from the BOE or configuration (for example Eway bill number/date and Remittance date/rate), THE Excel_Generator SHALL leave that field's cell empty with no characters, spaces, or placeholder text.
9. IF a source value was flagged as missing or unreadable during extraction, THEN THE Excel_Generator SHALL leave the corresponding cell empty with no characters, spaces, or placeholder text, rather than writing an inferred or placeholder value.

### Requirement 5: Populate the Excel line-item table from extracted data

**User Story:** As a User, I want each BOE line item written to the correct columns of the Item_Table, so that the output matches the sample layout.

#### Acceptance Criteria

1. THE Excel_Generator SHALL write Line_Item records into the Item_Table in ascending Item Serial Number order, beginning at the first data row immediately below the Item_Table header row (row 12), writing exactly one Line_Item record per row with no blank rows between records.
2. THE Excel_Generator SHALL write a sequential Sr. no. that equals 1 in the first Line_Item row and increases by exactly 1 in each subsequent Line_Item row, such that the Sr. no. values form the consecutive sequence 1 through N with no gaps, repeats, or skipped values for N Line_Items.
3. THE Excel_Generator SHALL write the extracted item description into the Description column.
4. THE Excel_Generator SHALL write the extracted CTH / HSN Code into the HSN CODE column.
5. THE Excel_Generator SHALL write the extracted Unit_Price_USD into the Unit Price in USD column.
6. THE Excel_Generator SHALL write the extracted quantity into the QTY column.
7. THE Excel_Generator SHALL write the extracted unit into the Unit column.
8. WHERE a Line_Item's unit, after trimming leading and trailing whitespace and ignoring letter case, equals "DOZ", THE Excel_Generator SHALL write the piece count (QTY multiplied by 12) into the pcs column.
9. WHERE a Line_Item's unit, after trimming leading and trailing whitespace and ignoring letter case, does not equal "DOZ", THE Excel_Generator SHALL leave the pcs column blank.
10. THE Excel_Generator SHALL write the extracted Assessable_Value into the CUSTOM ASS VALUE column.
11. IF a value required for any Item_Table cell of a Line_Item is missing or unreadable, THEN THE Excel_Generator SHALL leave only that cell blank, SHALL flag that cell as requiring User review, and SHALL write all remaining cells of that Line_Item's row without omitting the row.
12. THE Excel_Generator SHALL reproduce the Item_Table column order and each header label as a character-for-character exact match of the sample Item_Table, including spacing, punctuation, and letter case.

### Requirement 6: Compute derived per-line monetary values

**User Story:** As a User, I want duties, taxes, and landed costs calculated automatically per line, so that I do not compute them by hand.

#### Acceptance Criteria

1. THE Value_Calculator SHALL compute the per-line invoice Amount (USD) as Unit_Price_USD multiplied by QTY.
2. THE Value_Calculator SHALL compute the per-line purchase value in INR as the invoice Amount (USD) multiplied by the USD_Rate.
3. THE Value_Calculator SHALL compute the per-line SWS amount as the BCD amount multiplied by 0.10.
4. THE Value_Calculator SHALL compute the per-line total customs duty (excluding IGST) as the sum of the BCD amount and the SWS amount.
5. THE Value_Calculator SHALL compute the per-line IGST amount as the IGST rate, expressed as a decimal fraction (for example 0.18 for an 18% rate), multiplied by the sum of the Assessable_Value and the total customs duty (excluding IGST).
6. THE Value_Calculator SHALL compute the per-line combined duty as the sum of the total customs duty (excluding IGST) and the IGST amount.
7. THE Value_Calculator SHALL compute the per-line Land_Cost excluding GST as the sum of the purchase value in INR and the total customs duty (excluding IGST).
8. THE Value_Calculator SHALL compute the per-line Land_Cost including GST as the sum of the Land_Cost excluding GST and the IGST amount.
9. THE Value_Calculator SHALL compute the per-line purchase rate per unit as the Land_Cost excluding GST divided by QTY.
10. IF a Line_Item's QTY is zero, THEN THE Value_Calculator SHALL set the per-line purchase rate per unit to zero rather than performing a division by zero.
11. THE Excel_Generator SHALL write each computed value into its corresponding Item_Table column (Amount, Rate Per USD in purchase, SURCHARGE, TOTAL Custom Duty, GST, total custom duty, LAND COST OF PURCHASE WITHOUT GST, LAND COST OF PURCHASE WITH GST, Rate purchase per unit).
12. THE Value_Calculator SHALL retain every per-line computed value at full floating-point precision without rounding.
13. IF any input required to compute a per-line derived value (Unit_Price_USD, QTY, USD_Rate, BCD amount, IGST rate, or Assessable_Value) is missing or non-numeric, THEN THE Value_Calculator SHALL leave the dependent computed value blank and SHALL flag the affected Line_Item as requiring User review rather than substituting a default value.

### Requirement 7: Compute and write the totals row

**User Story:** As a User, I want a totals row that sums the columns, so that I can verify the workbook against the BOE invoice total.

#### Acceptance Criteria

1. THE Value_Calculator SHALL compute the total invoice Amount (USD) as the sum of all per-line invoice Amounts (USD).
2. THE Value_Calculator SHALL compute the column totals for Assessable_Value, total customs duty, IGST, and both Land_Cost columns as the sum of the corresponding per-line values.
3. WHEN no Line_Item records are present, THE Value_Calculator SHALL set every computed column total in the Totals_Row to zero.
4. THE Excel_Generator SHALL write the computed totals into the Totals_Row in the columns matching the sample layout, without altering their numeric precision.
5. THE Excel_Generator SHALL write the total package/CTN count, taken from the extracted BOE total number of packages, into the Totals_Row.
6. WHEN the Totals_Row is generated, IF the computed total invoice Amount (USD) and the BOE's declared total Invoice Amount differ by more than 0.01 USD, THEN THE Converter SHALL display a message to the User indicating the discrepancy and showing both the computed total and the declared total, and SHALL retain the generated workbook.
7. WHEN the Totals_Row is generated, IF the BOE's declared total Invoice Amount is absent or unreadable, THEN THE Converter SHALL display a message to the User indicating that the computed invoice total could not be verified against the BOE.

### Requirement 8: Preserve numeric fidelity and reproduce the exact output layout

**User Story:** As a User, I want the output workbook to match the sample format precisely with no rounding errors, so that downstream use and review are reliable.

#### Acceptance Criteria

1. THE Excel_Generator SHALL produce exactly one worksheet, named "Sheet1", and SHALL NOT add any additional worksheets to the output workbook.
2. THE Excel_Generator SHALL write the Header_Block at rows 1 through 8, the Item_Table header row at row 12, the line items in the rows immediately following the Item_Table header row, and the Totals_Row at the same row position as the sample workbook (row 61 in the sample), using the identical cell column positions of the sample workbook.
3. THE Excel_Generator SHALL reproduce the auxiliary section label text exactly as it appears in the sample ("DETAILS AS PER CHALLANS", "DETAILS AS PER TALLY", "CLEARANCE AND FORWARDING INVOICE", and the C&F detail block) at the same cell positions as the sample, and SHALL leave the data cells within those sections empty (containing no value).
4. THE Value_Calculator SHALL retain each value extracted directly from the BOE as its exact extracted numeric value, without applying rounding, truncation, or reformatting before it is written to the output.
5. THE Excel_Generator SHALL write each computed value to the output at the full numeric precision produced by the Value_Calculator, without rounding or truncation.
6. WHERE the sample workbook leaves a column empty for all line items (for example PARTY NAME, BILLING AMOUNT), THE Excel_Generator SHALL leave that column empty (containing no value and no whitespace) for all line items in the output.

### Requirement 9: Report parsing anomalies without dropping data

**User Story:** As a User, I want to be told about anything unusual in the BOE instead of having it silently omitted, so that I can trust the output's completeness.

#### Acceptance Criteria

1. WHEN the Converter completes a conversion, THE Converter SHALL present a summary stating the number of Line_Items extracted, the total invoice Amount in USD, and the count of fields flagged as requiring User review.
2. IF a Line_Item field cannot be parsed into its expected data type, THEN THE Converter SHALL write the raw extracted text for that field into the output without removing the field.
3. IF a Line_Item field cannot be parsed into its expected data type, THEN THE Converter SHALL mark that field with a visible indication that it requires User review.
4. IF a numeric value extracted from the BOE differs from the same value recomputed from its related fields by more than 0.01 in that value's unit, THEN THE Converter SHALL report both the extracted value and the recomputed value to the User.
5. THE Converter SHALL NOT remove a Line_Item from the output solely because one of its fields failed to parse.

### Requirement 10: Extract the BOE Part I duty-summary declared totals

**User Story:** As a User, I want the document-level duty totals read from the BOE's Part I duty summary, so that the Excel grand totals can be verified against the amounts the customs authority declared.

> Context: The existing PDF_Parser extracts header fields (Requirement 2) and per-line values (Requirement 3)
> but does NOT yet extract the Part I "C. DUTY SUMMARY" document-level totals. This requirement adds that
> extraction so Requirement 11 can compare the Excel grand totals against the BOE.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract the Declared_TOTAL_DUTY value from the Duty_Summary of the BOE.
2. THE PDF_Parser SHALL extract the Declared_IGST value from the Duty_Summary of the BOE.
3. THE PDF_Parser SHALL extract the Declared_BCD value from the Duty_Summary of the BOE.
4. THE PDF_Parser SHALL extract the Declared_SWS value from the Duty_Summary of the BOE.
5. THE PDF_Parser SHALL record each extracted Duty_Summary value as the value printed in the BOE, preserving its digits and decimal places without rounding, truncating, reformatting, or inferring content.
6. IF a Duty_Summary value required by criteria 1 through 4 cannot be located in the BOE, or its printed characters cannot be resolved into a numeric value, THEN THE Converter SHALL record that value as missing, SHALL report it to the User identified by its field name, and SHALL NOT substitute an inferred or default value.

### Requirement 11: Verify Excel grand totals against the BOE duty summary and highlight mismatches

**User Story:** As a User, I want the Excel grand totals for customs duty, GST, and combined duty checked against the BOE duty summary, and any mismatch highlighted in red, so that I can immediately see where the generated workbook disagrees with the BOE.

#### Acceptance Criteria

1. WHEN the Totals_Row is generated, THE Converter SHALL compare the column P Grand_Total (`Totals.total_customs_duty`) against the BOE declared customs duty excluding IGST, computed as Declared_BCD plus Declared_SWS.
2. WHEN the Totals_Row is generated, THE Converter SHALL compare the column Q Grand_Total (`Totals.total_igst`) against the Declared_IGST.
3. WHEN the Totals_Row is generated, THE Converter SHALL compare the column S Grand_Total (`Totals.total_customs_duty` plus `Totals.total_igst`) against the Declared_TOTAL_DUTY.
4. IF a compared Grand_Total differs from its corresponding BOE declared value by more than 1.00 INR, THEN THE Excel_Generator SHALL apply a solid red cell fill to that Grand_Total cell in the Totals_Row.
5. IF a compared Grand_Total differs from its corresponding BOE declared value by no more than 1.00 INR, THEN THE Excel_Generator SHALL write that Grand_Total cell without applying the red cell fill.
6. IF a compared Grand_Total differs from its corresponding BOE declared value by more than 1.00 INR, THEN THE Converter SHALL surface a Discrepancy to the User identifying the affected Excel column, the Excel Grand_Total, and the BOE declared value.
7. IF a BOE declared value required for a comparison in criteria 1 through 3 is absent or unreadable, including a missing Declared_BCD or Declared_SWS required for the column P comparison, THEN THE Converter SHALL report to the User that the corresponding Grand_Total could not be verified against the BOE, and SHALL write that Grand_Total cell without applying the red cell fill.
8. WHEN a Grand_Total mismatch is highlighted, THE Converter SHALL retain the generated workbook and SHALL make the workbook available for download.
9. WHEN the Converter presents the conversion summary, THE Converter SHALL include each grand-total verification Discrepancy in the summary reported to the User.

## Source-to-Target Field Mapping (informative)

This table records the field-by-field mapping derived from the sample PDF and Excel. It classifies each
target column as **Direct** (copied from the PDF), **Computed** (derived by the Value_Calculator), or
**Manual/External** (not available from the BOE alone — see Open Questions). The exact computed formulas
are stated normatively in Requirement 6.

| Target column (Item_Table) | Source classification | Notes |
|---|---|---|
| Sr. no. | Computed | Sequential 1..N |
| PARTY NAME | Manual/External | Blank for all line items in sample |
| BILLING AMOUNT | Manual/External | Blank for all line items in sample |
| AS PER TALLY NAME | Manual/External | Simplified item name; appears human-entered (see Open Q3) |
| Description | Direct | BOE item description (CTH item text) |
| HSN CODE | Direct | BOE CTH |
| CTN | Manual/External | Cartons per line; not clearly in BOE (see Open Q1) |
| QTY | Direct | BOE quantity (see Open Q2 on rounding/which qty field) |
| Unit | Direct | BOE unit (DOZ/KGS/NOS/PCS) |
| pcs | Computed | QTY×12 when Unit = DOZ, else blank |
| Unit Price in USD | Direct | BOE UPI |
| Amount | Computed | Unit Price × QTY |
| Rate Per USD in purchase | Computed | Amount × USD_Rate |
| CUSTOM ASS VALUE | Direct | BOE Assessable Value (INR) |
| LAND COST OF PURCHASE WITHOUT GST | Computed | Purchase INR + (BCD + SWS) |
| TOTAL Custom Duty | Computed | BCD + SWS |
| GST | Computed | IGST rate × (Assessable Value + BCD + SWS) |
| RATE OF DUTY IGST | Direct | BOE IGST rate |
| total custom duty | Computed | TOTAL Custom Duty + GST |
| RATE OF INTEREST (col T) | Direct | Actually the BCD rate (label is misleading) |
| CUST AIDC (col U) | Direct | Actually the BCD amount (label is misleading) |
| RATE OF INTEREST (col V) | Constant | SWS rate = 0.10 |
| SURCHARGE | Computed | BCD × 0.10 (SWS) |
| LAND COST OF PURCHASE WITH GST | Computed | Land cost without GST + GST |
| Rate purchase per unit | Computed | Land cost without GST ÷ QTY |

> Note: Several sample column headers are misleading (e.g., "CUST AIDC" holds the BCD amount; "RATE OF
> INTEREST" holds the BCD rate / SWS rate). The mapping above reflects the **actual** computed behavior
> observed in the sample, which the implementation must reproduce. This should be confirmed (Open Q4).

## Grand-Total Verification Mapping (informative)

This table records the working mapping used by Requirement 11 to verify Excel grand totals against the BOE
Duty_Summary. Values shown are from the sample BOE. The tolerance and highlight styling still need User
confirmation (see Open Q11-Q12).

| Excel Grand_Total (Totals_Row) | Computed as | BOE Duty_Summary value | Sample value |
|---|---|---|---|
| Column P — TOTAL Custom Duty | `Totals.total_customs_duty` | Declared_BCD + Declared_SWS | 184250.2 + 18810.2 = 203060.4 |
| Column Q — GST | `Totals.total_igst` | Declared_IGST | 258391 |
| Column S — total custom duty | `Totals.total_customs_duty + Totals.total_igst` | Declared_TOTAL_DUTY | 465302 |

> Note: Column P is verified against the BOE's declared BCD + SWS totals (per the resolution of Open Q10),
> which matches the modeled per-line customs duty basis (`total_customs_duty` = BCD + SWS). Because the
> verification basis and the modeled basis are the same duty components, the column P check no longer risks
> a false mismatch from unmodeled duty components (ACD, NCCD, ADD, CVD, G.CESS).

## Open Questions / Assumptions

These items were identified while reverse-engineering the sample and need User confirmation. Current
working assumptions are stated so the initial requirements remain actionable; answers may refine
Requirements 5-7.

1. **CTN (cartons per line item):** The per-line CTN values (e.g., 5, 45, 17) do not appear to come from
   the BOE line items; only the grand total (1357) matches the BOE package count. Are per-line CTN values
   taken from a separate packing list, entered manually, or derived by a rule? *Assumption:* manual/external for Milestone 1.
2. **QTY source and rounding:** The sample QTY appears to use a rounded quantity (e.g., 255 vs the BOE's
   255.15 KGS). Which BOE quantity field should drive QTY (commercial qty vs standard qty), and what
   rounding rule applies? *Assumption:* use the commercial quantity as printed, no rounding, pending confirmation.
3. **AS PER TALLY NAME:** This simplified name (with sample typos like "Decorative Itemsss") looks
   human-entered. Should the tool leave it blank for manual entry, or derive it from the description by a rule/lookup? *Assumption:* leave blank/manual for Milestone 1.
4. **Misleading column labels / duty model:** Confirm that "CUST AIDC" should hold the BCD amount and the
   two "RATE OF INTEREST" columns hold BCD rate and SWS rate (10%) respectively, and that AIDC/interest are
   not separately required. Confirm SWS is always 10% of BCD.
5. **USD_Rate origin:** Is the USD_Rate (95.3 in the sample) read from the BOE, supplied by the User per
   conversion, or read from a configuration value? *Assumption:* User-supplied per conversion.
6. **Company name:** Is the importer company name ("Gemini Unicom LLP") a fixed configuration value or
   read from the BOE? *Assumption:* configuration value.
7. **Rounding/precision of computed values:** Should computed monetary values be stored at full floating
   precision (as in the sample) or rounded to a fixed number of decimals? *Assumption:* full precision, matching the sample.
8. **Multi-invoice / multi-page BOEs:** The sample BOE has a single invoice. Should Milestone 1 support a
   BOE containing multiple invoices, or is single-invoice sufficient? *Assumption:* single-invoice sufficient for Milestone 1.
9. **Auxiliary sections (Challans, Tally, C&F):** Confirm these lower sections should be emitted as empty
   templates in Milestone 1 (data entry/Tally linkage deferred to Milestone B). *Assumption:* empty templates.
10. **Column P verification basis (Requirement 11): RESOLVED.** Per User decision, column P
    (`Totals.total_customs_duty`) is verified against the BOE's declared BCD + SWS totals
    (Declared_BCD + Declared_SWS), not against Declared_TOTAL_DUTY − Declared_IGST. This matches the modeled
    per-line customs duty basis (BCD + SWS), so the check does not require modeling the remaining BOE duty
    components (ACD, NCCD, ADD, CVD, G.CESS). Requirement 11 criterion 1 and the Grand-Total Verification
    Mapping have been updated accordingly.
11. **Verification tolerance (Requirement 11):** The BOE Duty_Summary prints values that appear rounded to
    whole rupees (e.g., TOTAL DUTY 465302), while the Excel grand totals are retained at full floating-point
    precision (Requirement 8). What INR tolerance should a "match" allow? *Assumption:* 1.00 INR, versus the
    0.01 USD tolerance used for the invoice-total cross-check in Requirement 7.
12. **Highlight styling (Requirement 11):** A solid red cell fill is assumed for mismatched grand-total cells.
    Confirm the exact fill shade (e.g., ARGB `FFFF0000`) and whether matching cells should ever receive an
    affirmative "verified" style (e.g., green) or simply remain unstyled. *Assumption:* solid red fill on
    mismatch only; matching cells left unstyled.
13. **BCD and SWS declared totals (Requirement 10): RESOLVED.** Following the resolution of Open Q10 toward a
    BCD + SWS basis for column P, Requirement 10 now extracts four declared Duty_Summary values —
    Declared_BCD, Declared_SWS, Declared_IGST, and Declared_TOTAL_DUTY — and Declared_BCD and Declared_SWS
    have been added to the Glossary.

## Milestone 1.1 — Field-Accuracy Corrections (bug-fix batch)

> These requirements capture defects found while converting real BOEs
> (`3. BE - 221981730062026INNSA1BE0040720261600.pdf` + `INV 1054.pdf`), verified against the
> client's own `bill_of_entry - with mistake.xlsx` (defects marked pink) and
> `bill_of_entry - corrected copy.xlsx` (the intended output). Where a correction contradicts an earlier
> assumption or clause it **supersedes** it; the superseded clause is named explicitly.
>
> Documentation gap noted: the code already contains a **Tally/JSON export** component
> (`tally_exporter.py`) plus per-line `other_duties_total` and buyer/seller extraction that the original
> Milestone 1 spec listed as out of scope ("Milestone B"). Requirements 18 below governs the JSON export
> defect only; a full spec pass documenting the Tally exporter is recommended as separate follow-up work.

### Requirement 12: Populate the USD invoice amount and USD rate in the Excel header from the BOE

**User Story:** As a User, I want the USD amount and USD rate shown in the Excel header block filled from the BOE, so that the header reflects the source document without manual entry.

> Supersedes Open Question 5 (USD_Rate User-supplied) and the "USD Amt" no-source treatment in
> Requirement 4.8. Evidence: corrected workbook cell `G3` = `13757.09` (the BOE invoice USD amount),
> `G2` = the USD rate.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract the total invoice amount expressed in USD from the BOE.
2. THE Excel_Generator SHALL write the extracted USD invoice amount into the header block "USD Amt" value cell (`G3`) at full numeric precision, without rounding or reformatting.
3. WHERE the BOE prints a USD exchange rate, THE PDF_Parser SHALL extract it and THE Excel_Generator SHALL write it into the "USD Rate" value cell (`G2`) at full numeric precision.
4. WHERE the BOE does not print a USD exchange rate, THE Converter SHALL fall back to the User-supplied USD rate as in Milestone 1.
5. IF the USD invoice amount cannot be located or resolved, THEN THE Converter SHALL record it as missing, report it to the User by field name, and leave the "USD Amt" cell blank rather than substituting a value.

### Requirement 13: Preserve BOE header values verbatim without date/number reformatting

**User Story:** As a User, I want dates and identifiers written exactly as printed on the BOE, so that no value is silently transformed.

> Reinforces Requirement 4.2/4.3 and Requirement 2.11. Evidence: mistake workbook wrote the invoice date
> as an Excel `datetime` (`2026-06-11`) where the intended value is the printed string `11-JUN-26` (cell
> `G4`).

#### Acceptance Criteria

1. THE Excel_Generator SHALL write each extracted header date (Invoice Date, BE Date, B/L Date) as the verbatim printed string, and SHALL NOT coerce it into an Excel date/serial value or otherwise reformat it.
2. THE Excel_Generator SHALL write each extracted header identifier (Invoice No, BE No, B/L No) as the verbatim printed value without reformatting.

### Requirement 14: Compute the `pcs` column for all piece-equivalent units (DOZ, GRS, THD)

**User Story:** As a User, I want the Excel `pcs` column to convert dozens, gross, and thousands to pieces just like the JSON export does, so that the two outputs agree.

> Supersedes Requirement 5.8/5.9 (DOZ-only). Evidence: corrected/mistake workbook `pcs` uses `=H*144`
> for a GRS-unit line. The JSON exporter already applies `DOZ×12, GRS×144, THD×1000`
> (`tally_exporter._UNIT_TO_PCS`); the Excel `pcs` rule must match it exactly.

#### Acceptance Criteria

1. WHERE a Line_Item's unit, trimmed and upper-cased, equals `"DOZ"`, THE Value_Calculator SHALL compute `pcs` as QTY × 12.
2. WHERE a Line_Item's unit, trimmed and upper-cased, equals `"GRS"`, THE Value_Calculator SHALL compute `pcs` as QTY × 144.
3. WHERE a Line_Item's unit, trimmed and upper-cased, equals `"THD"`, THE Value_Calculator SHALL compute `pcs` as QTY × 1000.
4. WHERE a Line_Item's unit is none of `DOZ`, `GRS`, `THD`, THE Value_Calculator SHALL leave the `pcs` cell blank.
5. THE Excel_Generator SHALL write the `pcs` cell (and, in formula mode, the multiplier in its formula) using the factor that matches the line's unit.
6. THE piece-conversion factors used by the Excel `pcs` column and by the JSON/Tally export SHALL be identical.

### Requirement 15: Use the BOE-declared SWS rate instead of a fixed 10%

**User Story:** As a User, I want the Social Welfare Surcharge to reflect what the BOE actually declares, so that an exempt (0%) line does not receive a fabricated 10% surcharge and cause a GST/grand-total mismatch.

> Supersedes Requirement 6.3 (SWS fixed at 0.10) and Open Question 4 ("Confirm SWS is always 10% of
> BCD"). Evidence: a real line carried a 0 SWS rate, yet the tool applied 10%, producing a GST mismatch.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract the per-line SWS rate and SWS amount as printed in the BOE, treating an exemption-driven 0 as the numeric value 0 (not blank).
2. THE Value_Calculator SHALL compute the per-line SWS amount from the BOE-declared SWS rate/amount rather than a hardcoded 10%.
3. WHERE the BOE-declared per-line SWS rate or amount is 0, THE Value_Calculator SHALL set the per-line SWS amount to 0.
4. WHERE the BOE-declared SWS rate/amount cannot be located for a line, THE Value_Calculator MAY fall back to 10% of the customs-duty base, and SHALL flag that line as requiring User review.
5. THE per-line total customs duty, IGST, combined duty, and land-cost values SHALL be derived from the corrected SWS amount so that the column P/Q/S grand totals reconcile with the BOE duty summary.

### Requirement 16: Retain full numeric precision for computed monetary values (no round-off)

**User Story:** As a User, I want computed values kept at full precision, so that totals reconcile and no rounding error accumulates.

> Reinforces Requirement 6.12 and 8.5; documents the observed defect. Evidence: mistake workbook `Q15` =
> `55.3` where the correct GST value is `55.29535`.

#### Acceptance Criteria

1. THE Value_Calculator SHALL retain every per-line computed monetary value (including the IGST/GST amount) at full floating-point precision, without rounding or truncation to a fixed number of decimals.
2. THE Excel_Generator SHALL write each directly-extracted value (including Unit Price in USD) at its full extracted precision, without rounding.
3. THE Converter SHALL source Unit Price in USD from the BOE Part II UPI field, not from any optional invoice attachment.
4. THE Converter SHALL confine any decimal rounding to display/number-format only, and SHALL NOT alter the stored cell value.

### Requirement 17: Grand-total mismatch highlight must fire on a genuine duty mismatch

**User Story:** As a User, I want the grand-total re-check to actually highlight a mismatched total in red, so that I can see disagreements with the BOE at a glance.

> Extends Requirement 11. Evidence: a real conversion had a duty/GST mismatch that was not highlighted,
> because the comparison basis diverged from the modeled totals (tied to the SWS defaulting defect,
> Requirement 15).

#### Acceptance Criteria

1. THE Converter SHALL compute the column P/Q/S grand-total verification on the same duty basis that produces the written totals, so a genuine mismatch beyond 1.00 INR is detected.
2. WHEN a grand-total mismatch beyond 1.00 INR is detected, THE Excel_Generator SHALL apply the solid red fill to the affected Totals_Row cell and the fill SHALL persist in the saved workbook (it SHALL NOT be overwritten by later writing steps).
3. WHEN the SWS correction (Requirement 15) removes a previously-fabricated surcharge, THE grand-total verification SHALL reflect the corrected totals.

### Requirement 18: NOS-unit line items must be carried into the JSON/Tally export

**User Story:** As a User, I want line items whose unit is NOS to appear in the JSON output, so that no product is dropped from the Tally import.

> Governs the JSON/Tally export defect only. Evidence: a NOS-unit line was absent from the produced JSON.

#### Acceptance Criteria

1. THE PDF_Parser SHALL extract the unit value `NOS` for any Line_Item that declares it, and SHALL NOT drop such a Line_Item.
2. THE JSON/Tally export SHALL include every Line_Item, including those whose unit is `NOS`, with its quantity and unit preserved (NOS is not a piece-equivalent unit, so its quantity is carried through unchanged).
3. THE Excel_Generator SHALL write `NOS` into the Unit column for such line items.
4. THE Converter SHALL NOT omit a Line_Item from either the Excel output or the JSON output solely because its unit is `NOS`.
