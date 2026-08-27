// ==========================================
// 1. ROUTING & ENTRY POINT
// ==========================================
function doGet(e) {
  e = e || {};
  const params = e.parameter || {};
  const page = params.page;

  if (page === 'digilocker_callback') {
    return handleDigiLockerCallback(e);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
      .setTitle('SIH Scheme Matcher')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 2. DATABASE & SETUP SERVICE (Google Sheets)
// ==========================================
function getDbId() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) {
    throw new Error("Database not set up. Please run setupProject() from the Apps Script editor first.");
  }
  return id;
}

function setupProject() {
  const ss = SpreadsheetApp.create("SIH_Unified_Database");
  PropertiesService.getScriptProperties().setProperty('DB_ID', ss.getId());

  const sheets = ["SCHEMES", "RULES", "RELATIONSHIPS", "DOCUMENTS"];
  sheets.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  Logger.log("Database Setup Complete. ID: " + ss.getId());
}

function seedDemoData() {
  const ss = SpreadsheetApp.openById(getDbId());
  
  // 1. SCHEMES (All 30 Verified Schemes: S1 to S30)
  const schemeSheet = ss.getSheetByName("SCHEMES");
  schemeSheet.clear();
  schemeSheet.appendRow(["scheme_id", "scheme_name", "target_group", "benefit", "application_url", "verified"]);
  
  schemeSheet.appendRow(["S1", "Pradhan Mantri MUDRA Yojana (PMMY)", "Micro enterprises / general", "Collateral-free loans up to ₹20 Lakhs", "jansamarth.in", "TRUE"]);
  schemeSheet.appendRow(["S2", "PM Employment Generation Programme (PMEGP)", "New micro-enterprises / individuals 18+", "Margin money subsidy up to ₹50 Lakhs (Mfg)", "kviconline.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S3", "Stand-Up India", "Women, SC/ST entrepreneurs", "Bank loans ₹10 Lakhs to ₹1 Crore (Greenfield)", "standupmitra.in", "TRUE"]);
  schemeSheet.appendRow(["S4", "Startup India / DPIIT Recognition", "Startups & innovators", "Tax exemptions, patent rebates, funding access", "startupindia.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S5", "Credit Guarantee Scheme for MSEs (CGTMSE)", "Micro and Small Enterprises", "Collateral-free credit guarantee up to ₹5 Crores", "cgtmse.in", "TRUE"]);
  schemeSheet.appendRow(["S6", "PM Formalisation of Micro Food Processing (PMFME)", "Micro food processors / SHGs", "35% capital subsidy up to ₹10 Lakhs", "pmfme.mofpi.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S7", "SFURTI", "Traditional artisans & clusters", "CFCs, raw material banks, skill training", "sfurti.msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S8", "National SC-ST Hub (NSSH)", "SC/ST owned MSEs", "Marketing assistance, capacity building, procurement", "scsthub.in", "TRUE"]);
  schemeSheet.appendRow(["S9", "Prime Minister's Scheme for Dev. of SCs (PMS-SC)", "Scheduled Caste beneficiaries", "Venture capital, concessional loans, skill support", "socialjustice.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S10", "National Scheduled Tribes Finance Corp (NSTFDC)", "Scheduled Tribe beneficiaries", "Concessional loans for income generation", "nstfdc.tribal.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S11", "National Safai Karamcharis Finance Corp (NSKFDC)", "Safai Karamcharis & dependents", "Concessional microfinance & training", "nskfdc.nic.in", "TRUE"]);
  schemeSheet.appendRow(["S12", "National Handicrafts Development Programme", "Handicraft artisans", "Design, technology upgrades, artisan cards", "handicrafts.nic.in", "TRUE"]);
  schemeSheet.appendRow(["S13", "National Handloom Development Programme", "Handloom weavers & groups", "Raw material supply, modern looms, marketing", "handlooms.nic.in", "TRUE"]);
  schemeSheet.appendRow(["S14", "PM Vishwakarma", "Traditional artisans (18 trades)", "Toolkit e-voucher ₹15k + Loan up to ₹3L at 5%", "pmvishwakarma.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S15", "PM Internship & Entrepreneurship Support", "Youth & early founders", "Stipends, experiential training & incubation", "mca.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S16", "ASPIRE (Promotion of Innovation & Rural Industries)", "Rural entrepreneurs & incubators", "Setting up Livelihood and Technology Business Incubators", "msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S17", "MSME Incubation Scheme", "Startups & innovators", "Financial assistance up to ₹15 Lakhs for idea phase & mentoring", "innovate.msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S18", "Procurement and Marketing Support (P&MS) Scheme", "Micro and Small Enterprises", "Subsidies for trade fairs, buyer-seller meets & marketing", "msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S19", "MSME Sustainable (ZED) Certification Scheme", "Manufacturing MSMEs", "Up to 80% subsidy on certification fees for quality & environment", "zed.msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S20", "Credit Linked Capital Subsidy & Tech Upgradation", "MSEs seeking machinery", "Up to 15% capital subsidy for technology upgrade", "msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S21", "MSE Cluster Development Programme (MSE-CDP)", "Artisan clusters & micro-enterprises", "Grants for Common Facility Centres (CFCs) & infrastructure", "msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S22", "Entrepreneurship Skill Development Programme (ESDP)", "Youth, women, SC/ST, traditional workers", "Subsidized skill training and entrepreneurship programs with stipends", "msme.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S23", "Coir Udyami Yojana (CUY)", "Coir unit entrepreneurs", "Credit-linked subsidy: 40% project subsidy up to ₹10 Lakhs", "coirboard.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S24", "Coir Vikas Yojana", "Coir artisans & units", "Support for skill development, R&D, and market promotion", "coirboard.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S25", "Skill Upgradation & Mahila Coir Yojana", "Women artisans in coir sector", "Specialized training & motorized spinning ratts with stipends", "coirboard.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S26", "National Handicrafts Development Programme (NHDP)", "Handicraft artisans", "Ambedkar Hastshilp Vikas Yojana, design & artisan credit cards", "handicrafts.nic.in", "TRUE"]);
  schemeSheet.appendRow(["S27", "National Handloom Development Programme (NHDP)", "Handloom weavers & groups", "Yarn supply subsidies, modern looms & shared infrastructure", "handlooms.nic.in", "TRUE"]);
  schemeSheet.appendRow(["S28", "MGNREGA Livelihood Asset Support", "Rural workers & smallholders", "Support for asset creation like livestock & rural livelihood sheds", "nrega.nic.in", "TRUE"]);
  schemeSheet.appendRow(["S29", "DAY-NRLM (SHG Enterprise Support)", "Women Self-Help Groups (SHGs)", "Revolving funds, community investment funds & bank linkages", "nrlm.gov.in", "TRUE"]);
  schemeSheet.appendRow(["S30", "PMFME Ecosystem & State-Linked Support", "Food processing micro-enterprises", "Seed capital, 35% capital subsidy up to ₹10 Lakhs & marketing", "pmfme.mofpi.gov.in", "TRUE"]);

  // 2. RULES
  const ruleSheet = ss.getSheetByName("RULES");
  ruleSheet.clear();
  ruleSheet.appendRow(["rule_id", "scheme_id", "field", "operator", "expected_value", "explanation"]);
  ruleSheet.appendRow(["R1", "S1", "age", ">=", "18", "Applicant must be 18 or older."]);
  ruleSheet.appendRow(["R2", "S2", "age", ">=", "18", "Must be 18 years or older to start a micro-enterprise."]);
  ruleSheet.appendRow(["R3", "S3", "gender", "==", "female", "Targeted at Women entrepreneurs (or SC/ST)."]);
  ruleSheet.appendRow(["R4", "S3", "funding_need", ">=", "1000000", "Minimum loan amount for Stand-Up India is ₹10 Lakhs."]);
  ruleSheet.appendRow(["R5", "S14", "occupation", "IN", "artisan,craftsman,bamboo,potter,weaver,carpenter", "Must practice one of the 18 specified traditional trades."]);
  ruleSheet.appendRow(["R6", "S6", "business_type", "IN", "food,processing,agro", "Must be a food processing enterprise."]);
  ruleSheet.appendRow(["R7", "S8", "category", "IN", "sc,st", "Must belong to SC or ST category."]);
  ruleSheet.appendRow(["R8", "S23", "occupation", "IN", "coir,artisan", "Must be engaged in coir production or processing."]);
  ruleSheet.appendRow(["R9", "S25", "gender", "==", "female", "Mahila Coir Yojana is exclusively for women artisans."]);
  ruleSheet.appendRow(["R10", "S29", "gender", "==", "female", "NRLM SHG enterprise support prioritizes women's self-help groups."]);

  // 3. DOCUMENTS
  const docSheet = ss.getSheetByName("DOCUMENTS");
  docSheet.clear();
  docSheet.appendRow(["scheme_id", "document_name"]);
  docSheet.appendRow(["S1", "Aadhaar"]);
  docSheet.appendRow(["S1", "Bank Statement"]);
  docSheet.appendRow(["S2", "Udyam Registration"]);
  docSheet.appendRow(["S2", "Project Report"]);
  docSheet.appendRow(["S3", "PAN"]);
  docSheet.appendRow(["S3", "Udyam Registration"]);
  docSheet.appendRow(["S3", "Category Certificate"]);
  docSheet.appendRow(["S14", "Aadhaar"]);
  docSheet.appendRow(["S14", "Bank Statement"]);
  docSheet.appendRow(["S14", "Artisan ID"]);
  docSheet.appendRow(["S19", "Udyam Registration"]);
  docSheet.appendRow(["S23", "Bank Statement"]);
  docSheet.appendRow(["S29", "SHG Membership Card"]);

  // 4. RELATIONSHIPS
  const relSheet = ss.getSheetByName("RELATIONSHIPS");
  relSheet.clear();
  relSheet.appendRow(["scheme_a", "relationship", "scheme_b"]);
  relSheet.appendRow(["S1", "INCOMPATIBLE", "S3"]);
  relSheet.appendRow(["S14", "COMPLEMENTARY", "S4"]);
  relSheet.appendRow(["S2", "COMPLEMENTARY", "S5"]);
  relSheet.appendRow(["S14", "COMPLEMENTARY", "S18"]);
  relSheet.appendRow(["S23", "COMPLEMENTARY", "S24"]);
  
  Logger.log("All 30 Schemes & Demo Data injected successfully.");
}

function getDatabase() {
  const ss = SpreadsheetApp.openById(getDbId());
  const db = {};
  const sheets = ["SCHEMES", "RULES", "RELATIONSHIPS", "DOCUMENTS"];
  
  sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      throw new Error('Sheet "' + name + '" is missing. Please run seedDemoData().');
    }
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    db[name] = data.map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  });
  return db;
}

// ==========================================
// 3. DETERMINISTIC LOGIC & MATCHING ENGINES
// ==========================================
const EligibilityEngine = {
  evaluate: (profile, db) => {
    const results = [];
    db.SCHEMES.forEach(scheme => {
      const rules = db.RULES.filter(r => r.scheme_id === scheme.scheme_id);
      let passed = [], failed = [], unknown = [];
      let isEligible = true;

      if (rules.length === 0) {
        results.push({ scheme: scheme, status: "POTENTIALLY_ELIGIBLE", passed: ["General eligibility criteria match."], failed: [], unknown: ["Requires detailed review."] });
        return;
      }

      rules.forEach(rule => {
        let actual = profile[rule.field];
        if (actual === undefined || actual === null || actual === "") {
          unknown.push(`Missing info: ${rule.field}`);
          isEligible = false;
          return;
        }

        let pass = false;
        const val = String(actual).toLowerCase();
        const expected = String(rule.expected_value).toLowerCase();

        if (rule.operator === '>=') pass = parseFloat(actual) >= parseFloat(expected);
        if (rule.operator === '<=') pass = parseFloat(actual) <= parseFloat(expected);
        if (rule.operator === '==') pass = val === expected;
        if (rule.operator === 'IN') pass = expected.split(',').map(s=>s.trim()).some(e => val.includes(e));

        if (pass) passed.push(rule.explanation);
        else { failed.push(rule.explanation); isEligible = false; }
      });

      let status = failed.length > 0 ? "NOT_ELIGIBLE" : (unknown.length > 0 ? "NEEDS_INFORMATION" : "ELIGIBLE");
      results.push({ scheme: scheme, status: status, passed: passed, failed: failed, unknown: unknown });
    });
    return results;
  }
};

const MatchingEngine = {
  rank: (eligibilityResults) => {
    const scores = { "ELIGIBLE": 3, "POTENTIALLY_ELIGIBLE": 2, "NEEDS_INFORMATION": 1, "NOT_ELIGIBLE": 0 };
    return eligibilityResults.sort((a, b) => scores[b.status] - scores[a.status]);
  }
};

const BundleEngine = {
  generate: (rankedSchemes, db) => {
    const eligible = rankedSchemes.filter(r => r.status === "ELIGIBLE" || r.status === "POTENTIALLY_ELIGIBLE");
    if (eligible.length === 0) return [];

    let primary = eligible[0].scheme;
    let bundle = [primary];

    eligible.slice(1).forEach(res => {
      const conflict = db.RELATIONSHIPS.find(rel =>
        (rel.scheme_a === primary.scheme_id && rel.scheme_b === res.scheme.scheme_id && rel.relationship === "INCOMPATIBLE") ||
        (rel.scheme_b === primary.scheme_id && rel.scheme_a === res.scheme.scheme_id && rel.relationship === "INCOMPATIBLE")
      );
      if (!conflict && bundle.length < 3) bundle.push(res.scheme);
    });

    return [{ name: "Recommended Growth Bundle", schemes: bundle }];
  }
};

const ReadinessEngine = {
  calculate: (rankedSchemes, profile, db) => {
    let readinessMap = {};
    const userDocs = profile.documents ? profile.documents.map(d => d.toLowerCase()) : [];

    rankedSchemes.forEach(res => {
      const schemeId = res.scheme.scheme_id;
      const reqDocs = db.DOCUMENTS.filter(d => d.scheme_id === schemeId).map(d => d.document_name);

      if (reqDocs.length === 0) {
        readinessMap[schemeId] = { target: res.scheme.scheme_name, score: 100, missing: [], required: [] };
        return;
      }

      let matchCount = 0;
      let missing = [];

      reqDocs.forEach(req => {
        if (userDocs.includes(req.toLowerCase())) {
          matchCount++;
        } else {
          missing.push(req);
        }
      });

      let score = Math.round((matchCount / reqDocs.length) * 100);
      readinessMap[schemeId] = {
        target: res.scheme.scheme_name,
        score: score,
        missing: missing,
        required: reqDocs
      };
    });

    return readinessMap;
  }
};

// ==========================================
// 4. GEMINI AI INTEGRATION (Profile Extraction)
// ==========================================
const GeminiService = {
  extractProfile: (text) => {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

    const fallbackParse = (t) => {
      let p = { documents: [] };
      const str = t.toLowerCase();
      if (str.includes("woman") || str.includes("female")) p.gender = "female";
      if (str.includes("rural")) p.rural = true;
      if (str.includes("karnataka")) p.state = "Karnataka";
      if (str.includes("bamboo") || str.includes("artisan") || str.includes("basket")) p.occupation = "artisan";

      if (str.includes("aadhaar") || str.includes("adhar")) p.documents.push("Aadhaar");
      if (str.includes("pan ")) p.documents.push("PAN");
      if (str.includes("bank statement")) p.documents.push("Bank Statement");
      if (str.includes("udyam")) p.documents.push("Udyam Registration");

      const match = str.match(/₹?(\d+,?\d+)/);
      if (match) p.funding_need = parseInt(match[1].replace(/,/g, ''));
      return p;
    };

    if (!apiKey) return fallbackParse(text);

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=${apiKey}`;
    const prompt = `Extract JSON. Return ONLY JSON, no markdown. Schema: {"gender":"female|male", "occupation":"string", "funding_need":number, "rural":boolean, "state":"string", "documents":["string"]}. Text: "${text}"`;

    try {
      const response = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json",
        payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
        muteHttpExceptions: true
      });
      const rawText = JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.replace(/```json|```/g, "");
      return JSON.parse(rawText);
    } catch (e) {
      Logger.log("Gemini failed, using fallback parser.");
      return fallbackParse(text);
    }
  }
};

// ==========================================
// 5. MAIN CONTROLLER API
// ==========================================
function processJourney(textInput, manualProfile) {
  try {
    let extracted = {};
    if (textInput) {
      extracted = GeminiService.extractProfile(textInput);
    }

    let profile = { ...extracted, ...(manualProfile || {}) };
    const extDocs = extracted.documents || [];
    const manDocs = (manualProfile && manualProfile.documents) ? manualProfile.documents : [];
    profile.documents = [...new Set([...extDocs, ...manDocs])];

    const db = getDatabase();
    const eligibilityResults = EligibilityEngine.evaluate(profile, db);
    const rankedSchemes = MatchingEngine.rank(eligibilityResults);
    const bundles = BundleEngine.generate(rankedSchemes, db);
    const readiness = ReadinessEngine.calculate(rankedSchemes, profile, db);

    return {
      success: true,
      profile: profile,
      rankedSchemes: rankedSchemes,
      bundles: bundles,
      readiness: readiness
    };

  } catch (e) {
    Logger.log("Error in processJourney: " + e.message);
    return { success: false, error: e.message };
  }
}

// ==========================================
// 6. DIGILOCKER INTEGRATION
// ==========================================
const DigiLockerService = {
  AUTH_URL: "https://digilocker.meripehchaan.gov.in/public/oauth2/1/authorize",
  TOKEN_URL: "https://digilocker.meripehchaan.gov.in/public/oauth2/1/token",
  DOCS_URL: "https://digilocker.meripehchaan.gov.in/public/oauth2/2/files/issued",

  DOCTYPE_MAP: {
    'ADHAR': 'Aadhaar',
    'PANCR': 'PAN',
    'CERTIFICATE': 'Category Certificate'
  },

  getAuthUrl: function () {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty('DIGILOCKER_CLIENT_ID');
    if (!clientId) {
      throw new Error("DigiLocker isn't configured yet. Add DIGILOCKER_CLIENT_ID and DIGILOCKER_CLIENT_SECRET under Project Settings > Script Properties.");
    }

    const state = Utilities.getUuid();
    CacheService.getScriptCache().put('dl_state_' + state, 'pending', 600);

    const redirectUri = ScriptApp.getService().getUrl();
    const params = {
      response_type: 'code',
      client_id: clientId,
      redirect_uri: redirectUri,
      state: state
    };
    const query = Object.keys(params)
      .map(k => k + '=' + encodeURIComponent(params[k]))
      .join('&');

    return this.AUTH_URL + '?' + query;
  },

  exchangeCodeForToken: function (code) {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty('DIGILOCKER_CLIENT_ID');
    const clientSecret = props.getProperty('DIGILOCKER_CLIENT_SECRET');
    const redirectUri = ScriptApp.getService().getUrl();

    const response = UrlFetchApp.fetch(this.TOKEN_URL, {
      method: 'post',
      payload: {
        grant_type: 'authorization_code',
        code: code,
        client_id: clientId,
        client_secret: clientSecret,
        redirect_uri: redirectUri
      },
      muteHttpExceptions: true
    });

    return JSON.parse(response.getContentText());
  },

  getIssuedDocuments: function (accessToken) {
    const response = UrlFetchApp.fetch(this.DOCS_URL, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    const data = JSON.parse(response.getContentText());
    const items = data.items || data.documents || [];
    const self = this;

    return items.map(function (doc) {
      return self.DOCTYPE_MAP[doc.doctype] || doc.name || doc.doctype;
    });
  }
};

function getDigiLockerAuthUrl() {
  return DigiLockerService.getAuthUrl();
}

function handleDigiLockerCallback(e) {
  const params = e.parameter || {};
  const code = params.code;
  const error = params.error;

  let payload;

  if (error) {
    payload = { source: 'ppDigiLockerBridge', type: 'error', message: error };
  } else if (!code) {
    payload = { source: 'ppDigiLockerBridge', type: 'error', message: 'No authorization code returned.' };
  } else {
    try {
      const tokenData = DigiLockerService.exchangeCodeForToken(code);
      if (tokenData.error) throw new Error(tokenData.error_description || tokenData.error);

      const docs = DigiLockerService.getIssuedDocuments(tokenData.access_token);
      payload = { source: 'ppDigiLockerBridge', type: 'documents', documents: docs };
    } catch (err) {
      payload = { source: 'ppDigiLockerBridge', type: 'error', message: err.message };
    }
  }

  const template = HtmlService.createTemplate(
    '<!DOCTYPE html><html><body style="font-family:sans-serif;text-align:center;padding:2rem;">' +
    '<p>DigiLocker verification complete. This window will close automatically.</p>' +
    '<script>' +
    '  var payload = <?!= payload ?>;' +
    '  if (window.opener) { window.opener.postMessage(payload, "*"); }' +
    '  setTimeout(function(){ window.close(); }, 800);' +
    '</script>' +
    '</body></html>'
  );
  template.payload = JSON.stringify(payload);

  return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 7. MULTILINGUAL AI CHATBOT (Gemini)
// ==========================================
const ChatbotService = {
  reply: function (history, message, lang) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return "The chatbot isn't fully set up yet — add a GEMINI_API_KEY in Script Properties.";
    }

    let db;
    try {
      db = getDatabase();
    } catch (e) {
      return "Chatbot error: Cannot read database. Run setupProject() first.";
    }

    // --- SANITIZE INPUTS (THIS IS THE FIX) ---
    const safeMessage = (message || '').toString().trim();
    
    const safeHistory = (history || []).filter(function(m) {
      return m && m.text && m.text.toString().trim().length > 0;
    }).map(function(m) {
      return {
        role: m.role === 'assistant' ? 'model' : 'user',
        parts: [{ text: m.text.toString().trim() }]
      };
    });
    
    // If somehow both are empty, bail out
    if (safeHistory.length === 0 && safeMessage.length === 0) {
      return "I didn't receive any message. Please type something and try again.";
    }
    // --- END SANITIZE ---

    const schemeSummary = db.SCHEMES
      .map(s => s.scheme_name + ': ' + s.benefit + ' (Target: ' + s.target_group + ')')
      .join('\n');

    const systemPrompt =
      'You are the support assistant for "project people", a navigator that helps Indian citizens ' +
      'find government support schemes. Answer using the scheme list below. If you are not sure, say so ' +
      'plainly and suggest the person use the official application_url or contact the relevant department. ' +
      'Keep answers short (2-4 sentences) and simple. Always reply in this language: ' + (lang || 'English') + '.\n\n' +
      'SCHEMES:\n' + schemeSummary;

    const contents = safeHistory;
    if (safeMessage.length > 0) {
      contents.push({ role: 'user', parts: [{ text: safeMessage }] });
    }

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=' + apiKey.trim();

    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ 
          contents: contents, 
          systemInstruction: { parts: [{ text: systemPrompt }] } 
        }),
        muteHttpExceptions: true
      });

      const responseText = response.getContentText();
      Logger.log('Gemini raw response: ' + responseText);

      const data = JSON.parse(responseText);

      if (data.error) {
        Logger.log('Gemini API error: ' + JSON.stringify(data.error));
        return "Chatbot Error: " + (data.error.message || JSON.stringify(data.error));
      }

      if (!data.candidates || !data.candidates[0] || !data.candidates[0].content || !data.candidates[0].content.parts) {
        Logger.log('Unexpected Gemini structure: ' + responseText);
        return "Sorry, I got an unexpected response. Please try again.";
      }

      return data.candidates[0].content.parts[0].text;

    } catch (e) {
      Logger.log('Chatbot catch error: ' + e.message);
      return "Sorry, I couldn't process that. Error: " + e.message;
    }
  }
}

function chatWithBot(history, message, lang) {
  try {
    return ChatbotService.reply(history, message, lang);
  } catch (e) {
    Logger.log('Chatbot system error: ' + e.message);
    return "Sorry, something went wrong: " + e.message;
  }
}

// ==========================================
// 8. DEBUG / TEST HELPERS
// ==========================================
function testGeminiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  Logger.log("Key exists: " + (apiKey ? "YES" : "NO"));
  Logger.log("Key length: " + (apiKey ? apiKey.length : 0));
  
  if (!apiKey) {
    Logger.log("❌ KEY IS MISSING!");
    return;
  }
  
  const cleanKey = apiKey.trim();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=' + cleanKey;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: 'Say hello in one word' }] }]
    }),
    muteHttpExceptions: true
  });
  
  const text = response.getContentText();
  Logger.log("HTTP Status: " + response.getResponseCode());
  
  const data = JSON.parse(text);
  if (data.error) {
    Logger.log("❌ GEMINI ERROR: " + JSON.stringify(data.error));
  } else if (data.candidates && data.candidates[0]) {
    Logger.log("✅ WORKING! Bot says: " + data.candidates[0].content.parts[0].text);
  }
}
