# AI-Based Fake Job Offer Detector

A browser-based prototype that helps identify suspicious job offer messages using text analysis, document hashing, domain checks, and cyber law awareness.

## Project Overview

This project detects possible fake job offers by checking offer letters or messages for scam indicators such as payment requests, urgency pressure, unrealistic salaries, unofficial email domains, suspicious links, and requests for sensitive personal information.

It is designed as an academic/project prototype and runs completely in the browser.

## Features

- Paste job offer text for analysis
- Upload `.txt` offer letters
- Detect suspicious scam keywords and patterns
- Extract emails, domains, and URLs from offer text
- Flag suspicious domains, short links, HTTP links, and payment URLs
- Compare company name with official domain
- Generate SHA-256 hash for uploaded documents
- Show risk score and verdict
- Save detection history in browser local storage
- Light and dark theme support
- Includes sample fake and genuine offer files

## Technologies Used

- HTML
- CSS
- JavaScript
- Browser Local Storage
- Web Crypto API for SHA-256 hashing

## Folder Structure

Crypto_P/
├── index.html
├── README.md
├── static/
│   ├── app.js
│   └── styles.css
├── samples/
│   └── fake_offer.txt
├── datasets/
│   ├── fake_offer_1.txt
│   ├── fake_offer_2.txt
│   └── genuine_offer_1.txt
├── scripts/
│   ├── create_presentation.py
│   ├── create_detailed_presentation.py
│   └── create_normal_presentation.py
└── presentation_assets/

## How to Run

Open the project folder:

cd "C:\Users\Harish\OneDrive\Desktop\Crypto_P"

Start a local server:

python -m http.server 8000

Open this URL in your browser:

http://localhost:8000

You can also open `index.html` directly in a browser.

## How to Use

1. Open the application.
2. Go to the **Fake Job Detection** page.
3. Enter the company name and official domain if available.
4. Paste the job offer text or upload a `.txt` offer letter.
5. Click the analyze button.
6. Review the risk score, verdict, findings, network checks, and document hash.
7. Check previous scans in the detection history page.

## Sample Input

Congratulations! You have been selected for a work from home job.
Pay a refundable security deposit of Rs. 2,000 using UPI within 24 hours.
No interview required. Contact us only on WhatsApp.

## Cyber Law Context

The project includes India-focused cyber law awareness references such as:

- National Cyber Crime Reporting Portal: https://cybercrime.gov.in/
- IT Act, 2000 Section 66C: Identity theft
- IT Act, 2000 Section 66D: Cheating by personation using computer resources

## Important Note

This project is only a prototype for academic and educational use. It does not replace official company verification, police investigation, legal advice, or cyber cell support.

## Author

Created for an academic cybersecurity/project demonstration.
