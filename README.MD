# Automated Legal Judgment Publisher

**Author:** Olayode Bolade

A robust automation system built with Google Apps Script designed to streamline the publication of legal judgments. This tool monitors a Google Drive folder, converts PDF/Word documents into machine-readable text, extracts key legal metadata (Ratio Decidendi, Subject Matters), and publishes the formatted content directly to a WordPress website using the REST API.

## Key Features

* **Format Conversion:** Automatically detects PDFs or Word docs and converts them to Google Docs to perform OCR/Text extraction.
* **Intelligent Parsing:**
    * Extracts the **Title Page** and Court hierarchy.
    * Identifies and formats the **Ratio Decidendi** using regex pattern matching.
    * Scrapes **Subject Matters** and sanitizes them for use as WordPress Tags.
* **Smart Categorization:** Automatically assigns WordPress categories (Supreme Court, Court of Appeal) based on file naming conventions.
* **PDF Embedding:** Generates an HTML embed code to display the original full judgment file within the WordPress post.
* **Duplicate Prevention:** Checks for existing tags before creating new ones to keep the database clean.
* **Cleanup:** Automatically manages storage by deleting temporary converted files after successful posting.

## Technology Stack

* **Language:** Google Apps Script (JavaScript)
* **Integrations:**
    * Google Drive API (File management & Conversion)
    * Google Docs API (Text parsing)
    * WordPress REST API (Content publishing & Tag management)

## Configuration

To use this script, open the `.gs` file and configure the following constants at the top of the script:

```javascript
const WP_POSTS_URL = "[https://your-domain.com/wp-json/wp/v2/posts](https://your-domain.com/wp-json/wp/v2/posts)";
const WP_TAGS_URL = "[https://your-domain.com/wp-json/wp/v2/tags](https://your-domain.com/wp-json/wp/v2/tags)";
const USERNAME = "your_wp_username";
const APP_PASSWORD = "your_application_password"; // Use an App Password, not your login password
const FOLDER_ID = "your_google_drive_folder_id";

Note: Update the Category IDs (CAT_JUDGMENTS, CAT_SC, etc.) to match the specific ID numbers in your WordPress installation.

## How It Works

* Ingest: The script scans the target Google Drive folder.
* Convert: If the file is a PDF, it converts it to a temporary Google Doc.
* Parse: It reads the text, looking for specific legal keywords (e.g., "RATIO DECIDENDI", "ISSUES").
* Tagging: It creates tags in WordPress dynamically based on the "Subject Matter" section of the document.
* Publish: It constructs an HTML payload with the parsed text and an iframe preview of the original file, then POSTs it to WordPress.

## License
* This project is created and maintained by Olayode Bolade.
