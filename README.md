# AI-Based Fake Job Offer Detector

A browser-based prototype that flags suspicious job offers by combining:

- AI-style text risk scoring
- Cryptographic document hashing
- Network security checks for email domains and URLs
- Cyber law reporting guidance

## How to Run

Open `index.html` directly in a modern browser.

For a local server:

```powershell
python -m http.server 8000
```

Then visit:

```text
http://localhost:8000
```

## Current Features

- Paste offer text or upload a `.txt` offer letter.
- Detects scam indicators such as security deposits, urgent pressure, unrealistic salary claims, and unofficial communication channels.
- Extracts emails, domains, and URLs from the offer.
- Flags suspicious domains, short links, HTTP links, payment URLs, and company/domain mismatches.
- Calculates SHA-256 hash for uploaded documents.
- Generates a risk score with clear reasons and legal next steps.
- Provides Home, Fake Job Detection, Detection History, and About screens.
- Saves recent detections in browser local storage.
- Supports light and dark themes.
- Includes `samples/fake_offer.txt` for quick testing.

## Cyber Law Context

For India-focused academic use, the app highlights:

- National Cyber Crime Reporting Portal: https://cybercrime.gov.in/
- IT Act, 2000 Section 66C: identity theft concepts.
- IT Act, 2000 Section 66D: cheating by personation using a computer resource.

Sources checked on 2026-05-14:

- National Government Services Portal listing for the National Cyber Crime Reporting Portal: https://services.india.gov.in/service/detail/national-cyber-crime-reporting-portal
- India Code PDF for the Information Technology Act, 2000: https://www.indiacode.nic.in/bitstream/123456789/1999/1/A2000-21%20%281%29.pdf

## Notes

This is a prototype for academic/project use. It does not replace professional legal advice, official company verification, or police/cyber cell investigation.
