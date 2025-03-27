# Webpage Scraper: Web2MS

# Overview

This Python script scrapes the main content section of webpages and converts them into MS Word (.docx) documents while maintaining the relative font sizes, structure, and formatting. It also extracts video transcripts if available.

## Features

- Reads a CSV file containing URLs and unique identifiers.
- Scrapes the main content section (text, images, and video transcripts).
- Preserves the font hierarchy (relative sizes of headings and paragraphs).
- Extracts and inserts video transcripts in the correct order.
- Saves each webpage as a Word document named using the unique identifier.
- Implements retry logic (2 attempts per URL in case of failure).
- Outputs documents in a configurable folder (`scraped_docs`).

## Prerequisites

Ensure you have Python installed on your system along with the required libraries. You can install the dependencies using:

```sh
pip install requests beautifulsoup4 python-docx
```

## Usage

1. **Prepare Input CSV**

   - The input CSV file should have the following format:
     ```csv
     URL,Identifier
     https://example.com/page1,document1
     https://example.com/page2,document2
     ```

2. **Run the Script**

   - Place your CSV file (e.g., `urls.csv`) in the same directory as the script.
   - Execute the script using:
     ```sh
     python webpage_scraper.py
     ```

3. **View Output**

   - The extracted Word documents will be saved in the `scraped_docs` directory.

## Configuration

- Modify the `OUTPUT_FOLDER` variable in the script to change the output directory.
- Adjust `RETRY_LIMIT` for more or fewer retries on failed requests.

## Error Handling

- The script logs errors and skips URLs that repeatedly fail after two attempts.
- If no content is found, the script will notify and move to the next URL.

## Limitations

- The script assumes content is inside the `<main>` tag or `<body>` if `<main>` is missing.
- Video transcripts must be present in a `track` tag for extraction.

## License

This project is open-source and can be modified as needed.

## Contact

For any issues or improvements, feel free to contribute!

