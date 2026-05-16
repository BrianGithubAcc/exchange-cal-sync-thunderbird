console.log("[ExCalSync] autodiscover.js loading");
/**
 * EWS Autodiscover module.
 * Discovers the EWS endpoint URL from an email address.
 * Implements the standard Autodiscover protocol with fallback URLs.
 */

// Depends on: ews/xml-builder.js (buildAutodiscoverRequest)

const AUTODISCOVER_PATHS = [
  // Standard autodiscover URLs tried in order
  "https://{domain}/autodiscover/autodiscover.xml",
  "https://autodiscover.{domain}/autodiscover/autodiscover.xml",
];

// Microsoft 365 autodiscover endpoint
const O365_AUTODISCOVER_URL =
  "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml";

/**
 * Extracts the domain from an email address.
 */
function getDomain(email) {
  const parts = email.split("@");
  if (parts.length !== 2 || !parts[1]) {
    throw new Error("Invalid email address format");
  }
  return parts[1].toLowerCase();
}

/**
 * Parses the Autodiscover XML response to extract the EWS URL.
 */
function parseAutodiscoverResponse(xmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "text/xml");

  // Check for parse errors
  const parseError = doc.querySelector("parsererror");
  if (parseError) {
    throw new Error("Invalid XML in autodiscover response");
  }

  // Look for EwsUrl in the Protocol section
  const protocols = doc.querySelectorAll("Protocol");
  for (const protocol of protocols) {
    const typeEl = protocol.querySelector("Type");
    if (typeEl && typeEl.textContent === "EXPR") {
      const ewsUrl = protocol.querySelector("EwsUrl");
      if (ewsUrl && ewsUrl.textContent) {
        return ewsUrl.textContent.trim();
      }
    }
  }

  // Fallback: check any EwsUrl element
  const ewsUrlEl = doc.querySelector("EwsUrl");
  if (ewsUrlEl && ewsUrlEl.textContent) {
    return ewsUrlEl.textContent.trim();
  }

  // Check for ASUrl (older format)
  const asUrlEl = doc.querySelector("ASUrl");
  if (asUrlEl && asUrlEl.textContent) {
    return asUrlEl.textContent.trim();
  }

  throw new Error("EWS URL not found in autodiscover response");
}

/**
 * Checks if a hostname resolves to a private/internal IP range.
 */
function isPrivateHostname(hostname) {
  // Block known private/reserved hostnames
  if (
    hostname === "localhost" ||
    hostname.endsWith(".local") ||
    hostname.endsWith(".internal") ||
    hostname.endsWith(".intranet")
  ) {
    return true;
  }

  // Check for IP address literals (IPv4 and IPv6)
  // IPv4 private ranges
  const ipv4Match = hostname.match(/^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/);
  if (ipv4Match) {
    const [, a, b] = ipv4Match.map(Number);
    if (
      a === 10 ||                          // 10.0.0.0/8
      (a === 172 && b >= 16 && b <= 31) || // 172.16.0.0/12
      (a === 192 && b === 168) ||          // 192.168.0.0/16
      a === 127 ||                          // 127.0.0.0/8 loopback
      (a === 169 && b === 254) ||          // 169.254.0.0/16 link-local
      a === 0                               // 0.0.0.0/8
    ) {
      return true;
    }
  }

  // IPv6 loopback and private
  const lower = hostname.toLowerCase();
  if (
    lower === "[::1]" ||
    lower.startsWith("[fe80:") ||    // link-local
    lower.startsWith("[fc") ||       // unique local
    lower.startsWith("[fd")          // unique local
  ) {
    return true;
  }

  return false;
}

/**
 * Validates an EWS URL to ensure it's a proper HTTPS URL pointing to a public host.
 */
function validateEwsUrl(url) {
  try {
    const parsed = new URL(url);
    if (parsed.protocol !== "https:") {
      throw new Error("EWS URL must use HTTPS");
    }
    if (!parsed.hostname || isPrivateHostname(parsed.hostname)) {
      throw new Error("EWS URL must not point to a private or internal host");
    }
    return parsed.href;
  } catch (e) {
    if (e.message.startsWith("EWS") || e.message.startsWith("Invalid")) {
      throw e;
    }
    throw new Error(`Invalid EWS URL: ${url}`);
  }
}

/**
 * Attempts Autodiscover against a single URL.
 */
async function tryAutodiscover(url, email, credentials) {
  const requestBody = buildAutodiscoverRequest(email);

  const headers = {
    "Content-Type": "text/xml; charset=utf-8",
  };

  if (credentials.type === "basic") {
    const encoded = btoa(
      String.fromCharCode(
        ...new TextEncoder().encode(
          `${credentials.username}:${credentials.password}`
        )
      )
    );
    headers["Authorization"] = `Basic ${encoded}`;
  } else if (credentials.type === "oauth") {
    headers["Authorization"] = `Bearer ${credentials.token}`;
  }

  const response = await fetch(url, {
    method: "POST",
    headers,
    body: requestBody,
    redirect: "follow",
  });

  if (!response.ok) {
    throw new Error(`Autodiscover HTTP ${response.status}: ${response.statusText}`);
  }

  const responseText = await response.text();
  return parseAutodiscoverResponse(responseText);
}

/**
 * Performs EWS Autodiscover to find the EWS endpoint.
 * Tries multiple URLs in order, including the O365 endpoint.
 *
 * @param {string} email - User's email address
 * @param {object} credentials - { type: 'basic'|'oauth', username?, password?, token? }
 * @returns {Promise<string>} The discovered EWS URL
 */
async function autodiscover(email, credentials) {
  const domain = getDomain(email);
  const errors = [];

  // Build list of URLs to try
  const urls = AUTODISCOVER_PATHS.map((tpl) =>
    tpl.replace("{domain}", domain)
  );

  // Add O365 endpoint for common Microsoft domains
  const microsoftDomains = [
    "outlook.com",
    "hotmail.com",
    "live.com",
    "microsoft.com",
  ];
  if (
    microsoftDomains.includes(domain) ||
    domain.endsWith(".onmicrosoft.com")
  ) {
    urls.unshift(O365_AUTODISCOVER_URL);
  } else {
    urls.push(O365_AUTODISCOVER_URL);
  }

  for (const url of urls) {
    try {
      console.log(`[EWS Autodiscover] Trying: ${url}`);
      const ewsUrl = await tryAutodiscover(url, email, credentials);
      const validated = validateEwsUrl(ewsUrl);
      console.log(`[EWS Autodiscover] Found EWS URL: ${validated}`);
      return validated;
    } catch (e) {
      console.log(`[EWS Autodiscover] Failed for ${url}: ${e.message}`);
      errors.push({ url, error: e.message });
    }
  }

  throw new Error(
    `Autodiscover failed for ${email}. Tried ${urls.length} endpoints. ` +
      `Errors: ${errors.map((e) => `${e.url}: ${e.error}`).join("; ")}`
  );
}

// Globals: autodiscover, getDomain, validateEwsUrl
