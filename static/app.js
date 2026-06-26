const suspiciousPatterns = [
  { regex: /\b(security deposit|registration fee|processing fee|training fee|refundable fee|verification fee)\b/i, weight: 22, reason: "Requests payment before joining." },
  { regex: /\b(upi|google pay|phonepe|paytm|bank transfer|account number|ifsc)\b/i, weight: 18, reason: "Contains personal payment instructions." },
  { regex: /\b(urgent|immediate|within 24 hours|limited seats|act fast|last chance)\b/i, weight: 12, reason: "Uses urgency pressure." },
  { regex: /\b(no interview|direct joining|guaranteed job|100% placement)\b/i, weight: 18, reason: "Promises hiring without normal selection steps." },
  { regex: /\b(passport|aadhaar|pan card|bank statement|otp|password)\b/i, weight: 14, reason: "Asks for sensitive identity or account information." },
  { regex: /\b(work from home.*earn|earn.*per day|salary.*lakh.*month)\b/i, weight: 14, reason: "Uses unusually attractive earning language." },
  { regex: /\b(whatsapp only|telegram|personal gmail|personal email)\b/i, weight: 12, reason: "Pushes communication away from official channels." },
  { regex: /\b(congratulation|congratulations).{0,50}\b(selected|shortlisted)\b/i, weight: 8, reason: "Instant selection wording detected." }
];

const shortenerDomains = new Set(["bit.ly", "tinyurl.com", "t.co", "goo.gl", "rebrand.ly", "cutt.ly", "is.gd"]);
const historyKey = "jobshieldDetectionHistory";
const themeKey = "jobshieldTheme";

const els = {
  navItems: document.querySelectorAll(".nav-item"),
  screens: document.querySelectorAll(".screen"),
  goButtons: document.querySelectorAll("[data-go]"),
  screenTitle: document.getElementById("screenTitle"),
  themeToggle: document.getElementById("themeToggle"),
  themeLabel: document.getElementById("themeLabel"),
  analyzeBtn: document.getElementById("analyzeBtn"),
  clearBtn: document.getElementById("clearBtn"),
  clearHistoryBtn: document.getElementById("clearHistoryBtn"),
  companyName: document.getElementById("companyName"),
  officialDomain: document.getElementById("officialDomain"),
  offerText: document.getElementById("offerText"),
  fileInput: document.getElementById("fileInput"),
  fileName: document.getElementById("fileName"),
  riskScore: document.getElementById("riskScore"),
  verdictBadge: document.getElementById("verdictBadge"),
  meterFill: document.getElementById("meterFill"),
  findingsList: document.getElementById("findingsList"),
  networkResults: document.getElementById("networkResults"),
  hashOutput: document.getElementById("hashOutput"),
  historyList: document.getElementById("historyList")
};

initializeTheme();
renderHistory();

els.navItems.forEach((item) => {
  item.addEventListener("click", () => showScreen(item.dataset.screen));
});

els.goButtons.forEach((button) => {
  button.addEventListener("click", () => showScreen(button.dataset.go));
});

els.themeToggle.addEventListener("click", () => {
  const nextTheme = document.body.dataset.theme === "dark" ? "light" : "dark";
  setTheme(nextTheme);
});

els.fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) {
    els.fileName.textContent = "No file selected";
    return;
  }

  els.fileName.textContent = file.name;
  els.offerText.value = await file.text();
});

els.analyzeBtn.addEventListener("click", async () => {
  const text = els.offerText.value.trim();
  const officialDomain = normalizeDomain(els.officialDomain.value);
  const result = analyzeOffer(text, officialDomain);
  const hash = await calculateHash(text);

  renderResults(result, hash);
  if (text) {
    saveHistory(result, hash);
    renderHistory();
  }
});

els.clearBtn.addEventListener("click", resetDetector);
els.clearHistoryBtn.addEventListener("click", () => {
  localStorage.removeItem(historyKey);
  renderHistory();
});

function showScreen(name) {
  const titleMap = {
    home: "Home",
    detector: "Fake Job Detection",
    history: "Detection History",
    about: "About"
  };

  els.navItems.forEach((item) => item.classList.toggle("active", item.dataset.screen === name));
  els.screens.forEach((screen) => screen.classList.toggle("active", screen.id === `${name}Screen`));
  els.screenTitle.textContent = titleMap[name] || "Home";
}

function initializeTheme() {
  const savedTheme = localStorage.getItem(themeKey) || "light";
  setTheme(savedTheme);
}

function setTheme(theme) {
  document.body.dataset.theme = theme;
  localStorage.setItem(themeKey, theme);
  els.themeToggle.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
  els.themeLabel.textContent = theme === "dark" ? "Dark" : "Light";
}

function analyzeOffer(text, officialDomain) {
  const findings = [];
  let score = 0;

  if (!text) {
    return {
      score: 0,
      verdict: "Waiting",
      level: "neutral",
      findings: ["Paste an offer and run analysis."],
      network: []
    };
  }

  for (const pattern of suspiciousPatterns) {
    if (pattern.regex.test(text)) {
      score += pattern.weight;
      findings.push(pattern.reason);
    }
  }

  const network = analyzeNetworkIndicators(text, officialDomain);
  score += network.reduce((sum, item) => sum + item.weight, 0);
  findings.push(...network.filter((item) => item.weight > 0).map((item) => item.reason));

  const wordCount = text.split(/\s+/).filter(Boolean).length;
  if (wordCount < 45) {
    score += 8;
    findings.push("Offer content is unusually short for a formal appointment letter.");
  }

  if (!/\b(interview|joining date|designation|ctc|human resources|hr department)\b/i.test(text)) {
    score += 8;
    findings.push("Missing several details normally present in legitimate offers.");
  }

  const normalizedScore = Math.min(100, score);
  return {
    score: normalizedScore,
    verdict: getVerdict(normalizedScore),
    level: getLevel(normalizedScore),
    findings: findings.length ? [...new Set(findings)] : ["No obvious scam indicators were found. Verify directly with the company before acting."],
    network
  };
}

function analyzeNetworkIndicators(text, officialDomain) {
  const items = [];
  const emails = extractMatches(text, /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g);
  const urls = extractMatches(text, /\bhttps?:\/\/[^\s<>"']+/gi);

  for (const email of emails) {
    const domain = normalizeDomain(email.split("@")[1]);
    if (isFreeMail(domain)) {
      items.push({ label: email, status: "bad", weight: 16, reason: `Recruiter email uses free-mail domain: ${domain}.` });
      continue;
    }

    if (officialDomain && !domain.endsWith(officialDomain)) {
      items.push({ label: email, status: "bad", weight: 14, reason: `Recruiter domain ${domain} does not match ${officialDomain}.` });
    } else {
      items.push({ label: email, status: "good", weight: 0, reason: `Email domain checked: ${domain}.` });
    }
  }

  for (const rawUrl of urls) {
    let parsed;
    try {
      parsed = new URL(rawUrl);
    } catch {
      items.push({ label: rawUrl, status: "bad", weight: 8, reason: "Malformed URL detected." });
      continue;
    }

    const domain = normalizeDomain(parsed.hostname);
    if (parsed.protocol !== "https:") {
      items.push({ label: rawUrl, status: "bad", weight: 10, reason: `Non-HTTPS link detected: ${domain}.` });
    }

    if (shortenerDomains.has(domain)) {
      items.push({ label: rawUrl, status: "bad", weight: 12, reason: `Shortened link detected: ${domain}.` });
    }

    if (/\b(pay|payment|upi|checkout|invoice|fee)\b/i.test(rawUrl)) {
      items.push({ label: rawUrl, status: "bad", weight: 15, reason: "Link appears related to payment or fees." });
    }

    if (officialDomain && !domain.endsWith(officialDomain) && !shortenerDomains.has(domain)) {
      items.push({ label: rawUrl, status: "bad", weight: 8, reason: `Link domain ${domain} does not match ${officialDomain}.` });
    }

    if (!items.some((item) => item.label === rawUrl)) {
      items.push({ label: rawUrl, status: "good", weight: 0, reason: `HTTPS link checked: ${domain}.` });
    }
  }

  if (!emails.length) {
    items.push({ label: "No recruiter email found", status: "bad", weight: 6, reason: "No recruiter email address found for domain verification." });
  }

  return items;
}

async function calculateHash(text) {
  if (!text) return "Not calculated";
  const data = new TextEncoder().encode(text);
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  return bufferToHex(hashBuffer);
}

function renderResults(result, hash) {
  els.riskScore.textContent = `${result.score}%`;
  els.verdictBadge.textContent = result.verdict;
  els.verdictBadge.className = `badge ${result.level}`;
  els.meterFill.style.width = `${result.score}%`;
  els.meterFill.style.backgroundColor = result.level === "danger" ? "var(--red)" : result.level === "warn" ? "var(--amber)" : "var(--accent)";

  els.findingsList.innerHTML = "";
  for (const finding of result.findings) {
    const li = document.createElement("li");
    li.textContent = finding;
    els.findingsList.appendChild(li);
  }

  els.networkResults.innerHTML = "";
  if (!result.network.length) {
    const pill = document.createElement("span");
    pill.className = "pill";
    pill.textContent = "No links checked yet";
    els.networkResults.appendChild(pill);
  } else {
    for (const item of result.network) {
      const pill = document.createElement("span");
      pill.className = `pill ${item.status}`;
      pill.textContent = item.label;
      pill.title = item.reason;
      els.networkResults.appendChild(pill);
    }
  }

  els.hashOutput.textContent = hash;
}

function saveHistory(result, hash) {
  const records = getHistory();
  const company = els.companyName.value.trim() || "Unknown company";
  const officialDomain = normalizeDomain(els.officialDomain.value) || "No domain";
  const preview = els.offerText.value.trim().slice(0, 120) || "No preview";
  records.unshift({
    company,
    officialDomain,
    preview,
    hash,
    score: result.score,
    verdict: result.verdict,
    level: result.level,
    createdAt: new Date().toLocaleString()
  });
  localStorage.setItem(historyKey, JSON.stringify(records.slice(0, 8)));
}

function renderHistory() {
  const records = getHistory();
  els.historyList.innerHTML = "";

  if (!records.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "No detections yet. Run an analysis to create a history entry.";
    els.historyList.appendChild(empty);
    return;
  }

  for (const record of records) {
    const card = document.createElement("article");
    card.className = "history-card";

    const detail = document.createElement("div");
    const title = document.createElement("strong");
    title.textContent = record.company;
    const meta = document.createElement("p");
    meta.textContent = `${record.createdAt} | ${record.officialDomain}`;
    const preview = document.createElement("p");
    preview.textContent = record.preview;
    const hash = document.createElement("code");
    hash.textContent = record.hash;

    detail.append(title, meta, preview, hash);

    const badge = document.createElement("span");
    badge.className = `badge ${record.level}`;
    badge.textContent = `${record.score}% ${record.verdict}`;

    card.append(detail, badge);
    els.historyList.appendChild(card);
  }
}

function resetDetector() {
  els.companyName.value = "";
  els.officialDomain.value = "";
  els.offerText.value = "";
  els.fileInput.value = "";
  els.fileName.textContent = "No file selected";
  renderResults(
    { score: 0, verdict: "Waiting", level: "neutral", findings: ["Paste an offer and run analysis."], network: [] },
    "Not calculated"
  );
}

function getHistory() {
  try {
    return JSON.parse(localStorage.getItem(historyKey)) || [];
  } catch {
    return [];
  }
}

function getVerdict(score) {
  if (score >= 70) return "Likely Fake";
  if (score >= 35) return "Suspicious";
  return "Lower Risk";
}

function getLevel(score) {
  if (score >= 70) return "danger";
  if (score >= 35) return "warn";
  return "safe";
}

function extractMatches(text, regex) {
  return [...new Set(text.match(regex) || [])];
}

function normalizeDomain(value) {
  return value
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/^www\./, "")
    .replace(/\/.*$/, "");
}

function isFreeMail(domain) {
  return ["gmail.com", "yahoo.com", "outlook.com", "hotmail.com", "proton.me", "protonmail.com", "rediffmail.com"].includes(domain);
}

function bufferToHex(buffer) {
  return [...new Uint8Array(buffer)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}
