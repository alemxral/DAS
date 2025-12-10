# Example Files

This directory contains example templates and data files to help you get started with the Document Automation System.

## Files Included

### Data File
- **example_data.xlsx**: Sample Excel file with customer data
  - Contains variables: `##name##`, `##email##`, `##amount##`, `##date##`
  - Includes 3 sample records

### Templates
- **example_template.docx**: Word template for generating customer letters
  - Uses all variables from the data file
  - Demonstrates basic text replacement

## How to Use

1. **Test the System**:
   - Open the web interface at `http://localhost:5000`
   - Click "New Job"
   - Select "File Path" for both template and data
   - Enter the full paths to these example files
   - Choose output formats (e.g., PDF, Word)
   - Click "Create Job"

2. **View Results**:
   - Wait for job to complete
   - Click "Download" to get ZIP with generated documents
   - Click "View" to preview files in browser

3. **Create Your Own**:
   - Use these examples as templates
   - Modify the Excel data structure
   - Update template placeholders
   - Add more variables as needed

## Excel Data Format

Your Excel file should have:
- **Row 1**: Variable names in `##variable##` format
- **Row 2+**: Data values for each record

Example:
```
| ##name##    | ##email##           | ##amount## | ##date##   |
|-------------|---------------------|------------|------------|
| John Smith  | john@example.com    | $1,000     | 2025-01-15 |
| Jane Doe    | jane@example.com    | $2,500     | 2025-01-20 |
```

## Template Format

Templates can contain any text with `##variable##` placeholders:

```
Dear ##name##,

Thank you for your order of ##amount##.
We will contact you at ##email## on ##date##.

Best regards,
Document Automation System
```

## Creating Excel Examples Manually

Since we cannot create binary Excel files directly, you'll need to create them manually:

1. **Create example_data.xlsx**:
   - Open Excel
   - In Row 1, enter: `##name##` | `##email##` | `##amount##` | `##date##`
   - In Row 2, enter: `John Smith` | `john@example.com` | `$1,000` | `2025-01-15`
   - In Row 3, enter: `Jane Doe` | `jane@example.com` | `$2,500` | `2025-01-20`
   - In Row 4, enter: `Bob Wilson` | `bob@example.com` | `$750` | `2025-01-25`
   - Save as `example_data.xlsx` in the `examples/` folder

2. **Create example_template.docx**:
   - Open Word
   - Type the following:
     ```
     CUSTOMER NOTIFICATION LETTER
     
     Date: ##date##
     
     Dear ##name##,
     
     This letter is to confirm your recent transaction.
     
     Transaction Details:
     - Customer Name: ##name##
     - Email Address: ##email##
     - Amount: ##amount##
     - Date: ##date##
     
     Thank you for your business. If you have any questions, please contact us at ##email##.
     
     Best regards,
     Document Automation Team
     ```
   - Save as `example_template.docx` in the `examples/` folder

## Tips

- Always use `##variable##` format (two hash marks on each side)
- Variable names should match between Excel first row and template
- Variables are case-sensitive
- You can use the same variable multiple times in a template
- Excel column order doesn't matter, only the variable names
