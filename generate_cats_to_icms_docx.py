from __future__ import annotations

from datetime import date

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def _add_field(paragraph, field_code: str) -> None:
    """
    Add a Word field code to a paragraph (e.g., TOC, PAGE).

    Word updates fields when the document is opened (or when user presses F9).
    """
    run = paragraph.add_run()
    r = run._r

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = field_code

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    r.append(fld_begin)
    r.append(instr)
    r.append(fld_sep)
    r.append(fld_end)


def _set_normal_style(doc: Document) -> None:
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)


def _add_footer_with_page_numbers(doc: Document) -> None:
    for section in doc.sections:
        footer = section.footer
        footer_par = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_field(footer_par, " PAGE ")


def _add_cover_page(doc: Document) -> None:
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("CATS to ICMS Integration System\nTechnical Documentation")
    r.bold = True
    r.font.size = Pt(26)

    doc.add_paragraph()

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = meta.add_run(
        "Repository: tampabaytimes-catstoicms-cc4024d115e5\n"
        "Technology: PHP\n"
        f"Document date: {date.today().strftime('%B %d, %Y')}"
    )
    r2.font.size = Pt(12)

    doc.add_paragraph()

    notice = doc.add_paragraph()
    notice.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = notice.add_run(
        "Internal technical reference. Configuration examples in this document are "
        "sanitized and do not contain production secrets."
    )
    r3.italic = True
    r3.font.size = Pt(10)

    # Page break
    doc.add_page_break()


def _add_toc(doc: Document) -> None:
    doc.add_paragraph("Table of Contents", style="Heading 1")
    p = doc.add_paragraph()
    _add_field(p, r'TOC \\o "1-3" \\h \\z \\u')
    doc.add_page_break()


def _add_code_block(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Consolas"
    run.font.size = Pt(9.5)


def _add_key_value_table(doc: Document, rows: list[tuple[str, str]], col_widths=(Inches(2.0), Inches(4.5))):
    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Parameter"
    hdr_cells[1].text = "Description / Example"
    for key, val in rows:
        row_cells = table.add_row().cells
        row_cells[0].text = key
        row_cells[1].text = val
    # Set widths (best-effort)
    for row in table.rows:
        row.cells[0].width = col_widths[0]
        row.cells[1].width = col_widths[1]
    doc.add_paragraph()


def build_docx(output_path: str) -> None:
    doc = Document()
    _set_normal_style(doc)

    # Reasonable margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    _add_cover_page(doc)
    _add_toc(doc)

    # 1. System Overview
    doc.add_paragraph("1. System Overview", style="Heading 1")
    doc.add_paragraph("Repository Information", style="Heading 2")
    doc.add_paragraph("Repository Name: tampabaytimes-catstoicms-cc4024d115e5")
    doc.add_paragraph("Technology Stack: PHP")
    doc.add_paragraph(
        "Primary Function: Automated data integration between external applicant tracking systems and the ICMS database"
    )
    doc.add_paragraph("Integration Points: CATS One API, HR Cloud API, SQL Server Database")
    doc.add_paragraph("Primary Server: PXLXPRDFCN01")
    doc.add_paragraph("Backup Server: PXLXPRDFCN02 (backup of PXLXPRDFCN01)")

    doc.add_paragraph("What This System Does", style="Heading 2")
    doc.add_paragraph(
        "This system is an automated data bridge that monitors external job application systems "
        "(CATS One and HR Cloud) for new Independent Contractor applications and imports them into "
        "the ICMS (Independent Contractor Management System) database. The objective is to eliminate "
        "manual data entry, enforce consistent formatting, and accelerate candidate processing."
    )

    # 2. Purpose and Business Context
    doc.add_paragraph("2. Purpose and Business Context", style="Heading 1")
    doc.add_paragraph("Business Problem", style="Heading 2")
    doc.add_paragraph(
        "Candidates applying for Independent Contractor or Newspaper Distributor positions submit their "
        "information through external systems. ICMS is the system of record for contractor lifecycle "
        "management; without automation, staff must manually transcribe applicant data, creating delays "
        "and increasing the likelihood of errors."
    )
    doc.add_paragraph("Business Solution", style="Heading 2")
    for item in [
        "Monitors external systems for new applications",
        "Extracts applicant information and application responses",
        "Validates and transforms data for ICMS compatibility",
        "Assigns distribution centers based on ZIP code mapping",
        "Creates ICMS prospect records and flags duplicates",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph("Business Value", style="Heading 2")
    for k, v in [
        ("Efficiency", "Removes manual entry effort for high-volume applicant flow"),
        ("Speed", "Applicants appear in ICMS shortly after submission"),
        ("Accuracy", "Standardized transformations reduce transcription errors"),
        ("Consistency", "Common formatting for imported records"),
        ("Scalability", "Supports growth without proportional staffing increases"),
    ]:
        p = doc.add_paragraph()
        r = p.add_run(f"{k}: ")
        r.bold = True
        p.add_run(v)

    # 3. System Architecture
    doc.add_paragraph("3. System Architecture", style="Heading 1")
    doc.add_paragraph("Server Infrastructure", style="Heading 2")
    doc.add_paragraph(
        "Primary Server (PXLXPRDFCN01) hosts integration scripts, configuration, and timestamp tracking files."
    )
    doc.add_paragraph(
        "Backup Server (PXLXPRDFCN02) maintains a synchronized copy and can be activated for failover."
    )

    doc.add_paragraph("High-Level Architecture", style="Heading 2")
    _add_code_block(
        doc,
        "\n".join(
            [
                "┌─────────────────┐         ┌──────────────────┐         ┌─────────────┐",
                "│   CATS One      │         │  Integration     │         │    ICMS      │",
                "│   API (v3)      │────────▶│  Scripts (PHP)   │────────▶│  Database    │",
                "│                 │         │  [PXLXPRDFCN01]  │         │  (SQL Server)│",
                "└─────────────────┘         └──────────────────┘         └─────────────┘",
                "                                    ▲",
                "┌─────────────────┐                │",
                "│   HR Cloud      │─────────────────┘",
                "│   API (v1)      │",
                "└─────────────────┘",
            ]
        ),
    )
    doc.add_paragraph()

    doc.add_paragraph("Component Structure", style="Heading 2")
    for title, desc in [
        ("CATS One Integration Script (cats2icms.php)", "Retrieves candidate applications from CATS One API v3."),
        ("HR Cloud Integration Script (hrcloud2icms.php)", "Retrieves applicant data from HR Cloud API v1."),
        ("Data Conversion Utilities (conversions.php)", "Formatting, normalization, mapping, and notification helpers."),
        ("Candidate Processing Module (process_candidate.php)", "Validations and extraction logic for application responses."),
    ]:
        p = doc.add_paragraph()
        r = p.add_run(f"{title}: ")
        r.bold = True
        p.add_run(desc)

    # 4. Integration Components
    doc.add_paragraph("4. Integration Components", style="Heading 1")

    # 4.1 CATS One Integration
    doc.add_paragraph("4.1 CATS One Integration", style="Heading 2")
    doc.add_paragraph("Overview", style="Heading 3")
    doc.add_paragraph(
        "The CATS One integration monitors CATS One for Independent Contractor applications and imports "
        "qualifying candidates into ICMS."
    )
    doc.add_paragraph("Key Features", style="Heading 3")
    for item in [
        "Incremental processing based on last run timestamp",
        "Job filtering for relevant Independent Contractor roles",
        "Custom field extraction for detailed application responses",
        "Activity review to confirm application relevance",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph("Process Flow (Summary)", style="Heading 3")
    for step in [
        "Read configuration and last processed timestamp",
        "Query CATS One API for candidates modified since last timestamp",
        "Filter candidates by activity/job title and exclude known test entries",
        "Fetch and map custom fields (paged)",
        "Validate, transform, and assign distribution center",
        "Insert ICMS prospect record and update timestamp",
    ]:
        doc.add_paragraph(step, style="List Number")

    doc.add_paragraph("Custom Field Mapping (Representative)", style="Heading 3")
    doc.add_paragraph(
        "The integration reads 20+ custom fields. The following list is a representative sample used by the import:"
    )
    for item in [
        "Personal information (first/last name, middle initial, business name)",
        "Contact information (address, city, state, ZIP, phones, emails)",
        "Qualifications (age, driver’s license, insurance, vehicle)",
        "PIC session preference and information source",
        "Prior distribution experience and availability",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    # 4.2 HR Cloud Integration
    doc.add_paragraph("4.2 HR Cloud Integration", style="Heading 2")
    doc.add_paragraph("Overview", style="Heading 3")
    doc.add_paragraph(
        "The HR Cloud integration monitors HR Cloud for Newspaper Distributor applications and imports "
        "qualifying applicants into ICMS."
    )
    doc.add_paragraph("Key Features", style="Heading 3")
    for item in [
        'Job title validation for roles starting with "Newspaper Distributor"',
        "Answer extraction from JSON response payloads",
        "Incremental processing using xLastUpdated timestamps",
        "Optional debug mode for troubleshooting",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_paragraph("Process Flow (Summary)", style="Heading 3")
    for step in [
        "Read configuration and last processed timestamp",
        "Query HR Cloud API for applicants updated since last timestamp (paged)",
        "Filter by job title and parse answers",
        "Validate and normalize field values",
        "Assign distribution center and check recent duplicates",
        "Insert ICMS prospect record and advance timestamp",
    ]:
        doc.add_paragraph(step, style="List Number")

    # 5. Data Flow and Processes
    doc.add_paragraph("5. Data Flow and Processes", style="Heading 1")
    doc.add_paragraph("Complete Data Flow", style="Heading 2")
    _add_code_block(
        doc,
        "\n".join(
            [
                "External System (CATS One / HR Cloud)",
                "    │",
                "    │ API Request (authenticated)",
                "    ▼",
                "Integration Script (PHP)",
                "    │",
                "    │ Extract candidate/applicant payload",
                "    ▼",
                "Transformation Layer",
                "    │ • Format phone numbers",
                "    │ • Normalize state codes",
                "    │ • Map ZIP → distribution center",
                "    │ • Translate coded values",
                "    ▼",
                "Validation Layer",
                "    │ • Required fields present",
                "    │ • Address completeness",
                "    │ • Duplicate detection (recent history)",
                "    ▼",
                "ICMS Database (SQL Server) → ICMS Application",
            ]
        ),
    )
    doc.add_paragraph()

    doc.add_paragraph("Transformation Rules", style="Heading 2")
    doc.add_paragraph("Phone Number Formatting", style="Heading 3")
    doc.add_paragraph(
        "Inputs may include country codes, spaces, or punctuation. The integration removes non-digits and "
        "formats values consistently as XXX-XXX-XXXX when possible."
    )
    doc.add_paragraph("State Code Normalization", style="Heading 3")
    doc.add_paragraph(
        "State values are normalized to a two-letter uppercase code. Invalid or missing values default to FL."
    )
    doc.add_paragraph("Distribution Center Assignment", style="Heading 3")
    doc.add_paragraph(
        "ZIP codes are mapped to predefined distribution center zones. ZIP values that do not map to a zone "
        "result in an unassigned center and trigger an email notification to the applicant."
    )
    doc.add_paragraph("Information Source Translation", style="Heading 3")
    doc.add_paragraph("Numeric IDs from the source system are mapped to a standardized source name.")
    doc.add_paragraph("PIC Session Mapping", style="Heading 3")
    doc.add_paragraph("Numeric IDs are mapped to a session city name for ICMS routing and reporting.")

    doc.add_paragraph("Duplicate Detection", style="Heading 2")
    for item in [
        "The integration checks for previous applications within the last 30 days.",
        "Matching is performed using normalized phone numbers.",
        'PrevApplicant is set to "P" when a match is found; otherwise "N".',
    ]:
        doc.add_paragraph(item, style="List Bullet")

    # 6. Configuration and Setup
    doc.add_paragraph("6. Configuration and Setup", style="Heading 1")
    doc.add_paragraph("Server Deployment", style="Heading 2")
    doc.add_paragraph(
        "Scripts run on PXLXPRDFCN01 under a scheduled task. PXLXPRDFCN02 maintains a synchronized copy "
        "for failover and recovery."
    )
    doc.add_paragraph("Configuration Files", style="Heading 2")
    doc.add_paragraph("Production: cats2icms.config")
    doc.add_paragraph("Development: cats2icmsdev.config")

    doc.add_paragraph("Configuration Parameters (Sanitized)", style="Heading 2")
    _add_key_value_table(
        doc,
        rows=[
            ("token", "<REDACTED> (CATS One API token)"),
            ("tz", 'America/New_York'),
            ("dbserver", "busdev-sql-1.ad.sptimes.com"),
            ("dbuser", "ICMSUser"),
            ("dbpass", "<REDACTED>"),
            ("dbname", "tpcICMS"),
            ("hrcloud_key", "<REDACTED>"),
            ("hrcloud_secret", "<REDACTED>"),
            ("debug", '"true" or "false" (HR Cloud only)'),
        ],
    )

    doc.add_paragraph("Timestamp Tracking Files", style="Heading 2")
    for item in [
        "catsLastReadTime.txt: last processed timestamp (production).",
        "catsLastReadTimedev.txt: last processed timestamp (development).",
        "Timestamps are stored as ISO 8601 datetimes and updated after successful runs.",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph("Distribution Center ZIP Code Mappings", style="Heading 2")
    doc.add_paragraph(
        "The system uses predefined ZIP code ranges for each distribution center. Mapping updates should be "
        "reviewed periodically and synchronized to both servers."
    )
    for dc, zips in [
        ("Distribution Center 2 (Plant/Tampa Area)", "33701-33715, 33716, 33760, 33762, 33772-33774, 33776-33778, 33781-33782, 33785-33786"),
        ("Distribution Center 4 (Largo)", "33756, 33760, 33762, 33764, 33767, 33770-33771, 33772-33774, 33776-33778, 33781-33782, 33785-33786"),
        ("Distribution Center 6 (Palm Harbor)", "33755-33756, 33759, 33761, 33763-33765, 33767, 33770-33771, 34677, 34681, 34683-34685, 34688-34689, 34695, 34698"),
        ("Distribution Center 7 (West Pasco)", "34610, 34652-34655, 34667-34668, 34690-34691"),
        ("Distribution Center 8 (Hernando)", "34601-34602, 34604, 34606-34609, 34613-34614"),
        ("Distribution Center 9 (South Hillsborough)", "33510-33511, 33527, 33534, 33547, 33563, 33565-33567, 33569-33570, 33572-33573, 33578-33579, 33584, 33592, 33594, 33596, 33598, 33602, 33605, 33610, 33617, 33619, 33637"),
        ("Distribution Center 10 (Central Pasco)", "33525, 33543-33545, 33548-33549, 33559, 33576, 33647, 34637-34639, 33523, 33535, 33537, 33540-33542"),
        ("Distribution Center 12 (West Tampa)", "33556, 33558, 33603-33604, 33606-33607, 33609, 33611-33616, 33618, 33624-33626, 33629, 33634-33635"),
        ("Distribution Center 0 (Unassigned)", "Any ZIP not matching the above zones triggers an email notification."),
    ]:
        p = doc.add_paragraph()
        r = p.add_run(f"{dc}: ")
        r.bold = True
        p.add_run(zips)

    # 7. Features and Capabilities
    doc.add_paragraph("7. Features and Capabilities", style="Heading 1")
    doc.add_paragraph("Automated Data Import", style="Heading 2")
    for item in [
        "Regular polling of external systems for new/updated records",
        "Incremental processing to avoid rework",
        "Batch handling for throughput",
        "Continues processing when individual records fail validation",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_paragraph("Data Validation", style="Heading 2")
    for item in [
        "Required field checks (name, address components, phone)",
        "Address completeness validation",
        "Basic format checks prior to insertion",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_paragraph("Data Transformation", style="Heading 2")
    for item in [
        "Phone standardization",
        "State normalization",
        "Coded value translation (PIC session, information source)",
        "ZIP-to-distribution-center assignment",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_paragraph("Duplicate Management", style="Heading 2")
    for item in [
        "30-day lookback for previous applications",
        "Phone-based matching",
        "PrevApplicant flag for staff visibility",
    ]:
        doc.add_paragraph(item, style="List Bullet")
    doc.add_paragraph("Notification System", style="Heading 2")
    for item in [
        "Email notification for unmatched ZIP codes",
        "Includes distribution center options and meeting schedule guidance",
        "SMTP delivery via mailx (server-configured)",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    # 8. Error Handling and Notifications
    doc.add_paragraph("8. Error Handling and Notifications", style="Heading 1")
    doc.add_paragraph("Error Types and Handling", style="Heading 2")
    for heading, bullets in [
        ("Database Connection Errors", ["Detected via connection exceptions.", "Logged and processing stops; requires intervention."]),
        ("API Connection Errors", ["Failed requests/timeouts are logged.", "Processing continues and will retry on next run."]),
        ("Data Validation Errors", ["Invalid records are skipped with context logged.", "Records can be reviewed and corrected manually."]),
        ("Rate Limiting", ["Handles throttling by pacing requests in controlled batches."]),
    ]:
        doc.add_paragraph(heading, style="Heading 3")
        for b in bullets:
            doc.add_paragraph(b, style="List Bullet")

    doc.add_paragraph("Email Notifications for Unmatched ZIP Codes", style="Heading 2")
    for item in [
        "Trigger: candidate ZIP does not match a configured distribution center zone.",
        "Action: send an informational email to the candidate explaining options and next steps.",
        "Delivery: mailx with SMTP configuration; a brief delay is used to reduce mail server load.",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph("Logging and Monitoring", style="Heading 2")
    for item in [
        "Console output provides progress and error context for each run.",
        "Timestamp files provide checkpointing for incremental processing and recovery.",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    # 9. Maintenance and Operations
    doc.add_paragraph("9. Maintenance and Operations", style="Heading 1")
    doc.add_paragraph("Scheduled Execution", style="Heading 2")
    doc.add_paragraph(
        "The integration is designed to run on a schedule (cron or a task scheduler). The active schedule "
        "should be verified directly on PXLXPRDFCN01."
    )
    doc.add_paragraph("Execution Commands", style="Heading 3")
    _add_code_block(
        doc,
        "\n".join(
            [
                "php cats2icms.php",
                "php hrcloud2icms.php",
                "php test_email.php <name> <email>",
            ]
        ),
    )
    doc.add_paragraph()

    doc.add_paragraph("Monitoring Points", style="Heading 2")
    for item in [
        "Timestamp files updated after successful runs",
        "Console output indicates number of processed candidates/applicants",
        "New prospect records appear in the ICMS database",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph("Maintenance Tasks", style="Heading 2")
    for item in [
        "Weekly: review errors and anomalies from recent runs.",
        "Monthly: verify ZIP mapping accuracy and synchronize updates to the backup server.",
        "Quarterly: validate custom field mappings against source system changes.",
        "As needed: rotate credentials and update configuration on both servers.",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph("Backup and Recovery", style="Heading 2")
    for item in [
        "PXLXPRDFCN02 maintains a synchronized copy of scripts and configuration.",
        "Timestamp files allow controlled reprocessing by adjusting the last-run value.",
        "Failover requires enabling scheduling on the backup server and confirming network/database access.",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_paragraph("Security Considerations", style="Heading 2")
    for item in [
        "Restrict file permissions on configuration files; treat API tokens and DB credentials as secrets.",
        "Use secure transport for API and database connections where available.",
        "Avoid embedding production secrets in documentation; use sanitized examples instead.",
    ]:
        doc.add_paragraph(item, style="List Bullet")

    # Appendix
    doc.add_paragraph("Appendix", style="Heading 1")
    doc.add_paragraph("Custom Field ID Reference", style="Heading 2")
    doc.add_paragraph("CATS One Custom Fields", style="Heading 3")
    for line in [
        "249826: Acknowledgement",
        "249886: Information Source",
        "249829: Application Date",
        "249832: First Name",
        "249835: Last Name",
        "249383: Middle Initial",
        "249841: Business Name",
        "249844: Title",
        "249847: Attention",
        "249850: Address",
        "249853: City",
        "249856: State",
        "249589: ZIP Code",
        "249862: Phone Home",
        "249865: Phone Cell",
        "249868: Phone Fax",
        "249871: Email Primary",
        "249874: Email Alternate",
        "249877: At Least 18",
        "249880: Valid FL License",
        "249883: PIC Session",
        "249889: Valid Car Insurance",
        "249892: Vehicle",
        "249885: Has Distributed Papers",
        "249898: Distributed Where",
        "249901: Early Mornings OK",
    ]:
        doc.add_paragraph(line, style="List Bullet")

    doc.add_paragraph("PIC Session ID Mapping", style="Heading 2")
    for line in [
        "645334: St. Pete",
        "645337: Hernando",
        "645340: Largo",
        "645343: Palm Harbor",
        "645346: S. Hillsborough",
        "645349: W. Tampa",
    ]:
        doc.add_paragraph(line, style="List Bullet")

    doc.add_paragraph("Information Source ID Mapping", style="Heading 2")
    for line in [
        "645364: craigslist",
        "654523: expressemployment",
        "654520: hiregy",
        "645379: indeed",
        "645385: otherPublication",
        "645388: otherWebsite",
        "645382: radio",
        "645373: referralFromFriend",
        "645376: referralFromIC",
        "645370: tearOff",
        "645367: flyer",
        "645352: timesNewsAd",
        "645355: timesSignAd",
        "645358: timesWeb",
        "645361: tbt",
    ]:
        doc.add_paragraph(line, style="List Bullet")

    doc.add_paragraph("Database Table Structure", style="Heading 2")
    doc.add_paragraph("ProspectContractor Table Fields", style="Heading 3")
    fields = [
        "ProspectID (Primary Key)",
        "FirstName",
        "LastName",
        "MiddleInit",
        "BusinessName",
        "SAddress",
        "POBox",
        "City",
        "State",
        "Zip",
        "PersonalPhone",
        "EmailAdd",
        "LocDate",
        "IsPic",
        "HasPicAgreement",
        "SSNEIN",
        "DLLID",
        "DLLExp",
        "DOB",
        "DistributionCenter_DCID",
        "Comments",
        "IsIC",
        "SessionCity",
        "InfoSource",
        "DeliveryType",
        "PrevApplicant",
    ]
    for f in fields:
        doc.add_paragraph(f, style="List Bullet")

    doc.add_paragraph("Document Version History", style="Heading 2")
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Version"
    hdr[1].text = "Date"
    hdr[2].text = "Author"
    hdr[3].text = "Changes"
    row = table.add_row().cells
    row[0].text = "1.0"
    row[1].text = "2024-01-XX"
    row[2].text = "System Analysis"
    row[3].text = "Initial documentation"

    # Footer: page numbers
    _add_footer_with_page_numbers(doc)

    # Ensure the document is a single continuous section (python-docx creates one by default).
    # If needed later: doc.add_section(WD_SECTION.NEW_PAGE)
    _ = WD_SECTION  # silence unused import in some linters

    doc.save(output_path)


if __name__ == "__main__":
    build_docx("/workspace/CATS_to_ICMS_Integration_System_Technical_Documentation.docx")

