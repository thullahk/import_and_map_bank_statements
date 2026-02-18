# import_and_map_bank_statements
Bank Statement Import &amp; Smart Mapping – Odoo 16 Community Overview  This module enhances the default bank statement import functionality in Odoo 16 Community by providing advanced statement import, intelligent field mapping, and flexible configuration options.  It simplifies the process of importing bank statements from various formats.

# Key Features
1. Universal File Support
Multiple Formats: Supports importing bank statements from both CSV and XLSX (Excel) files.

Sheet Selection: For Excel files with multiple sheets, users can select the specific sheet to import.

Header Detection: Includes an option to use the first row of the file as a header for easier column identification.

2. Flexible Column Mapping
Dynamic Mapping: Users can map file columns to specific Odoo fields (e.g., Date, Label/Reference, Partner, Amount, and Currency).

Example Preview: Displays "Example Content" from the uploaded file during the mapping process to help users ensure they are selecting the correct columns.

3. Advanced Formatting Options
Encoding Support: Allows users to specify file encoding (e.g., UTF-8, Windows-1252, Latin1) to prevent character corruption.

Custom Separators: Supports various CSV delimiters such as commas, semicolons, and tabs.

Value Formatting: Users can define date formats (e.g., %Y-%m-%d) and specify decimal or thousand separators for numeric values.

4. Intelligent Partner & Currency Handling
Automated Partner Creation: Can be configured to automatically create new partners in Odoo if the partner name in the statement does not exist.

Foreign Currency Support: Capable of handling transactions in foreign currencies by mapping "Foreign Currency Code" and "Foreign Currency Amount" fields.

5. User Interface & Integration
Dashboard Integration: Adds an "Import Statement" button directly to the Odoo Accounting Dashboard (Kanban view) for bank journals.

Test Import: Features a "Test Import" button that allows users to validate their data and mapping before final processing to avoid errors.

Error Handling: Includes a configurable "On Error" behavior (e.g., skip or fail) to manage data inconsistencies during import.
