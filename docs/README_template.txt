Deduplication and Official Linkage Engine (DOLE) v1.0.0 - DOLE {{PROVINCE}} Provincial Office
Programmed by A. Enage (aenage@gmail.com)

A Windows desktop tool to detect duplicates, link user records to master databases, and flag government officials in beneficiary lists. Includes a dedicated Auditor tool.

FEATURES
- High-accuracy name matching with nickname equivalence, phonetics, and fuzzy ratios
- Match classes: Exact, Fuzzy, No Match
- Generates a formatted Excel report (Dashboard, User File Data, Analysis Report)
- Optional PDF export if Microsoft Excel is installed
- Standalone Auditor GUI for suspicious matches review

REQUIREMENTS
- 64-bit Windows 8.1, 10, or 11 (or Windows Server 2012 R2+)
- No installation required. Run the .exe from any folder or a USB flash drive.
- Optional: Microsoft Excel (desktop) if you want the report exported as PDF.

USAGE (MAIN APP)
- Run `DOLE_v1.0.0_{{PROVINCE}}.exe` to launch the application.
- Select your user data file (CSV, XLSX, or TXT) and province profile
- Click Start to run analysis; progress and logs are shown
- Output report is saved next to your input as: <input>_<Province>_report_<N>.xlsx
- If Microsoft Excel is installed, you can export the report to PDF; otherwise an Excel (.xlsx) report is produced

DATA HANDLING NOTES
- Contact Number is not normalized; preserved verbatim for follow-up calls

AUDITOR TOOL
- Launch the DOLE_v1.0.0_{{PROVINCE}}_Auditor.exe.
- Select a generated Analysis Report (.xlsx) and an output CSV path.
- The tool flags pairs with issues (examples):
  - Birthdate/Sex/Suffix mismatch.
  - Weak first/last name similarity; middle initial mismatch.
  - City mismatch when only names are available.
  - “Exact” remark with unexpectedly low overall similarity.
- Output is `suspicious_matches_<report base>.csv` for quick review.

TROUBLESHOOTING
- First run may take up to a minute while Windows prepares required components. Subsequent runs are faster.
- You can run the EXE directly from a local folder or external drive; no admin rights needed.
- PDF export failed: Ensure Microsoft Excel (desktop) is installed and licensed
- Column mapping issues: Ensure the first row has clear column headers (First Name, Middle Name, Last Name, Birthdate, City, Sex, Contact Number)

SUPPORT
- If you need help, please contact your DOLE MIMAROPA administrator.

VERSION INFORMATION
- Version: 1.0.0
- Release Date: {{CURRENT_DATE}}
- Copyright © 2025 DOLE - All Rights Reserved

