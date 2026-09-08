const SHEET_NAME = "record";
// Where applicants can check their status — update this if the page ever moves.
const STATUS_PAGE_URL = "https://spadesatlums.com/status.html";
// Admin password is stored only in Apps Script Script Properties.
const ADMIN_KEY_PROPERTY = "ADMIN_KEY";
const ADMIN_TOKEN_TTL_SECONDS = 1800;
const HEADERS = [
"candidateId",
"fullName",
"lumsEmail",
"phone",
"school",
"studentStatus",
"gender",
"previousEvents",
"startedAt",
"completedAt",
"firstPreference",
"secondPreference",
"standardScore",
"firstPreferenceScore",
"secondPreferenceScore",
"totalScore",
"maxScore",
"status",
"responses"
];
// ========================================
// GET
// ========================================
function doGet(e) {
  try {
    const params = e.parameter || {};

    // Simple deployment/sheet health check (no applicant data exposed).
    if (params.action === "health") {
      const sheet = getSheet_();
      ensureHeaders_(sheet);
      return json_({
        success: true,
        sheet: SHEET_NAME,
        rows: Math.max(0, sheet.getLastRow() - 1)
      });
    }

    // ====================================
    // ADMIN LIST
    // ====================================
    if (params.action === "list") {
      if (!isValidAdminToken_(params.adminToken)) {
        return json_({ success: false, error: "Unauthorized" });
      }

      return json_({
        success: true,
        applicants: getApplicants_()
      });
    }

    // ====================================
    // PUBLIC STATUS LOOKUP
    // Requires BOTH candidate ID + LUMS email.
    // Returns only minimum public information.
    // ====================================
    const candidateId = String(params.candidateId || "").trim();
    const email = String(params.email || "").trim().toLowerCase();

    if (!candidateId || !email) {
      return json_({
        found: false,
        error: "Candidate ID and LUMS email are required"
      });
    }

    if (!isValidCandidateId_(candidateId) || !isValidLumsEmail_(email)) {
      return json_({ found: false, error: "Application not found" });
    }

    const sheet = getSheet_();
    ensureHeaders_(sheet);
    const values = sheet.getDataRange().getValues();

    if (values.length < 2) {
      return json_({ found: false, error: "Application not found" });
    }

    const headers = values[0];
    const candidateIdCol = headers.indexOf("candidateId");
    const emailCol = headers.indexOf("lumsEmail");

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const rowCandidateId = candidateIdCol >= 0
        ? String(row[candidateIdCol] || "").trim()
        : "";

      const rowEmail = emailCol >= 0
        ? String(row[emailCol] || "").trim().toLowerCase()
        : "";

      if (rowCandidateId === candidateId && rowEmail === email) {
        const applicant = rowToObject_(headers, row);

        return json_({
          found: true,
          data: {
            candidateId: applicant.candidateId || "",
            status: applicant.status || "Submitted",
            firstPreference: applicant.firstPreference || "",
            secondPreference: applicant.secondPreference || ""
          }
        });
      }
    }

    return json_({ found: false, error: "Application not found" });

  } catch (err) {
    return json_({ found: false, error: "Server error" });
  }
}

// ========================================
// POST
// ========================================
function doPost(e) {
  let data = {};
  try {
    data = parsePostData_(e);

    // ====================================
    // ADMIN LOGIN
    // ====================================
    if (data.action === "adminLogin") {
      const suppliedPassword = String(data.password || "");
      const adminKey = getAdminKey_();

      if (!adminKey || !secureCompare_(suppliedPassword, adminKey)) {
        return json_({ success: false, error: "Invalid credentials" });
      }

      const token = Utilities.getUuid();
      CacheService.getScriptCache().put(
        "adminToken:" + token,
        "1",
        ADMIN_TOKEN_TTL_SECONDS
      );

      return json_({
        success: true,
        adminToken: token,
        expiresIn: ADMIN_TOKEN_TTL_SECONDS
      });
    }

    // ====================================
    // ADMIN LOGOUT
    // ====================================
    if (data.action === "adminLogout") {
      const token = String(data.adminToken || "");
      if (token) {
        CacheService.getScriptCache().remove("adminToken:" + token);
      }
      return json_({ success: true });
    }

    // ====================================
    // ADMIN STATUS UPDATE
    // ====================================
    if (data.action === "updateStatus") {
      if (!isValidAdminToken_(data.adminToken)) {
        return json_({ success: false, error: "Unauthorized" });
      }

      return updateStatus_(data);
    }

    // ====================================
    // APPLICATION SUBMISSION
    // ====================================
    return appendApplication_(data);

  } catch (err) {
    if (data && data._fromForm) return submissionResponse_({ success: false, error: "Server error" });
    return json_({ success: false, error: "Server error" });
  }
}

// ========================================
// SAVE APPLICATION
// ========================================
function appendApplication_(data) {
  const lock = LockService.getScriptLock();
  const t0 = Date.now();

  try {
    lock.waitLock(15000);
    const tLock = Date.now();

    const validation = validateApplication_(data);
    if (!validation.valid) {
      return data._fromForm ? submissionResponse_({ success: false, error: validation.error }) : json_({ success: false, error: validation.error });
    }
    const tValidate = Date.now();

    const sheet = getSheet_();
    ensureHeaders_(sheet);
    const headers = getHeaders_(sheet);
    const tSheet = Date.now();
    // Duplicate email submissions are allowed.
    const email = String(data.lumsEmail).trim().toLowerCase();

    // Use the browser candidate ID only as a display/lookup identifier after
    // validating its format and uniqueness. If it is missing/invalid/collides,
    // generate a fresh server-side ID.
    let candidateId = String(data.candidateId || "").trim().toUpperCase();
    if (!isValidCandidateId_(candidateId) || candidateIdExists_(sheet, headers, candidateId)) {
      candidateId = generateCandidateId_(sheet, headers);
    }
    const tCandidateId = Date.now();
    const row = new Array(headers.length).fill("");

    setValue_(row, headers, "candidateId", candidateId);
    setValue_(row, headers, "fullName", cleanText_(data.fullName, 100));
    setValue_(row, headers, "lumsEmail", email);
    setValue_(row, headers, "phone", cleanText_(data.phone, 30));
    setValue_(row, headers, "school", cleanText_(data.school, 100));
    setValue_(row, headers, "studentStatus", cleanText_(data.studentStatus, 50));
    setValue_(row, headers, "gender", cleanText_(data.gender, 30));
    setValue_(row, headers, "previousEvents", cleanText_(data.previousEvents, 2000));
    setValue_(row, headers, "startedAt", cleanText_(data.startedAt, 100));
    setValue_(row, headers, "completedAt", new Date());
    setValue_(row, headers, "firstPreference", data.firstPreference);
    setValue_(row, headers, "secondPreference", data.secondPreference);

    // IMPORTANT: scores are calculated here, never trusted from the browser.
    const scores = calculateScores_(data.responses);
    setValue_(row, headers, "standardScore", scores.standardScore);
    setValue_(row, headers, "firstPreferenceScore", scores.firstPreferenceScore);
    setValue_(row, headers, "secondPreferenceScore", scores.secondPreferenceScore);
    setValue_(row, headers, "totalScore", scores.totalScore);
    setValue_(row, headers, "maxScore", scores.maxScore);
    const tScore = Date.now();

    setValue_(row, headers, "status", "Submitted");
    setValue_(row, headers, "responses", JSON.stringify(normalizeResponses_(data.responses)));

    sheet.appendRow(row);
    const tAppend = Date.now();

    console.log(
      "SPADES_TIMING lockWait=" + (tLock - t0) +
      "ms validate=" + (tValidate - tLock) +
      "ms sheetSetup=" + (tSheet - tValidate) +
      "ms candidateId=" + (tCandidateId - tSheet) +
      "ms scoreAndBuildRow=" + (tScore - tCandidateId) +
      "ms append=" + (tAppend - tScore) +
      "ms TOTAL=" + (tAppend - t0) + "ms"
    );

    sendConfirmationEmail_({
      email: email,
      fullName: cleanText_(data.fullName, 100),
      candidateId: candidateId
    });

    return data._fromForm ? submissionResponse_({
      success: true, found: true, candidateId: candidateId, message: "Application submitted successfully"
    }) : json_({
      success: true, found: true, candidateId: candidateId, message: "Application submitted successfully"
    });

  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// ========================================
// FORM SUBMISSION RESPONSE
// ========================================
function submissionResponse_(result) {
  const html = '<!doctype html><html><body><script>' +
    'parent.postMessage({type:"SPADES_SUBMISSION_RESULT",success:' +
    (result.success ? 'true' : 'false') + ',candidateId:' + JSON.stringify(result.candidateId || "") +
    ',error:' + JSON.stringify(result.error || "") + '},"*");' +
    '</script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========================================
// CONFIRMATION EMAIL
// ========================================
function sendConfirmationEmail_(info) {
  try {
    const statusLink = STATUS_PAGE_URL + "?candidateId=" + encodeURIComponent(info.candidateId);
    MailApp.sendEmail({
      to: info.email,
      subject: "SPADES 2026 — Application Received (" + info.candidateId + ")",
      body:
        "Hi " + (info.fullName || "there") + ",\n\n" +
        "Thanks for applying to SPADES 2026! Your application has been received.\n\n" +
        "Your Candidate ID: " + info.candidateId + "\n\n" +
        "Check your application status anytime here:\n" + statusLink + "\n\n" +
        "You'll need this Candidate ID along with the LUMS email you applied with to look it up.\n\n" +
        "— SPADES"
    });
  } catch (mailErr) {
    // Never let a mail failure affect a submission that already saved successfully.
    console.log("SPADES_MAIL_ERROR " + (mailErr && mailErr.message ? mailErr.message : mailErr));
  }
}

// ========================================
// VALIDATION
// ========================================
function validateApplication_(data) {
  const email = String(data.lumsEmail || "").trim().toLowerCase();

  if (!String(data.fullName || "").trim()) return { valid: false, error: "Full name is required" };
  if (!isValidLumsEmail_(email)) return { valid: false, error: "A valid LUMS email is required" };
  if (!String(data.phone || "").trim()) return { valid: false, error: "Phone is required" };

  const school = String(data.school || "").trim();
  const studentStatus = String(data.studentStatus || "").trim();
  const gender = String(data.gender || "").trim();
  const first = String(data.firstPreference || "").trim();
  const second = String(data.secondPreference || "").trim();

  const allowedSchools = ["SBASSE", "MGSHSS", "SAHSOL", "SDSB"];
  const allowedStatuses = ["Freshman", "Sophomore", "Junior", "Senior"];
  const allowedGenders = ["Male", "Female", "Prefer not to say"];

  if (allowedSchools.indexOf(school) === -1) return { valid: false, error: "Invalid school" };
  if (allowedStatuses.indexOf(studentStatus) === -1) return { valid: false, error: "Invalid student status" };
  if (allowedGenders.indexOf(gender) === -1) return { valid: false, error: "Invalid gender" };

  if (!first || !second || first === second) {
    return { valid: false, error: "Two distinct department preferences are required" };
  }

  if (!Array.isArray(data.responses)) {
    return { valid: false, error: "Invalid responses" };
  }

  if (data.responses.length !== 11) {
    return { valid: false, error: "Invalid response count" };
  }

  for (let i = 0; i < data.responses.length; i++) {
    const r = data.responses[i];
    if (!r || typeof r !== "object") return { valid: false, error: "Invalid response data" };
    if (!String(r.question || "").trim() || !String(r.selectedAnswer || "").trim()) {
      return { valid: false, error: "Incomplete responses" };
    }
  }

  return { valid: true };
}

function isValidLumsEmail_(email) {
  return /^[A-Za-z0-9._%+-]+@lums\.edu\.pk$/i.test(email);
}

function isValidCandidateId_(id) {
  return /^SP-26-[A-Z0-9]{6}$/.test(String(id || ""));
}

function cleanText_(value, maxLength) {
  return String(value || "").trim().slice(0, maxLength);
}

// ========================================
// UPDATE STATUS
// ========================================
function updateStatus_(data) {
const sheet = getSheet_();
ensureHeaders_(sheet);
const values =
sheet.getDataRange().getValues();
if (values.length < 2) {


return json_({
  success: false,
  error: "No applications found"
});


}
const headers = values[0];
const candidateIdCol =
headers.indexOf("candidateId");
const statusCol =
headers.indexOf("status");
if (
candidateIdCol === -1 ||
statusCol === -1
) {


return json_({
  success: false,
  error: "Required columns are missing"
});


}
const candidateId =
String(
data.candidateId || ""
).trim();
const newStatus =
String(
data.status || ""
).trim();
if (!candidateId || !newStatus) {


return json_({
  success: false,
  error:
    "Candidate ID and status are required"
});


}
for (
let i = 1;
i < values.length;
i++
) {


const rowCandidateId =
  String(
    values[i][candidateIdCol] || ""
  ).trim();


if (
  rowCandidateId === candidateId
) {

  sheet
    .getRange(
      i + 1,
      statusCol + 1
    )
    .setValue(newStatus);


  return json_({

    success: true,

    found: true,

    message:
      "Status updated successfully"

  });

}


}
return json_({
success: false,
found: false,
error: "Candidate not found"
});
}

// ========================================
// GET APPLICANTS
// ========================================
function getApplicants_() {
const sheet = getSheet_();
ensureHeaders_(sheet);
const values =
sheet.getDataRange().getValues();
if (values.length < 2) {
return [];
}
const headers = values[0];
return values
.slice(1)
.filter(function(row) {


  return row.some(
    function(cell) {

      return String(
        cell || ""
      ).trim() !== "";

    }
  );

})
.map(function(row) {

  return rowToObject_(
    headers,
    row
  );

});


}

// ========================================
// GET SHEET
// ========================================
function getSheet_() {
const ss =
SpreadsheetApp
.getActiveSpreadsheet();
let sheet =
ss.getSheetByName(
SHEET_NAME
);
if (!sheet) {


sheet =
  ss.insertSheet(
    SHEET_NAME
  );


}
return sheet;
}

// ========================================
// ENSURE HEADERS
// ========================================
function ensureHeaders_(sheet) {
const lastColumn =
sheet.getLastColumn();
// Completely empty sheet
if (lastColumn === 0) {


sheet
  .getRange(
    1,
    1,
    1,
    HEADERS.length
  )
  .setValues([HEADERS]);

return;


}
const existingHeaders =
sheet
.getRange(
1,
1,
1,
lastColumn
)
.getValues()[0]
.map(function(h) {


    return String(
      h || ""
    ).trim();

  });


HEADERS.forEach(
function(header) {


  if (
    !existingHeaders
      .includes(header)
  ) {

    sheet
      .getRange(
        1,
        sheet.getLastColumn() + 1
      )
      .setValue(header);

    existingHeaders.push(
      header
    );

  }

}


);
}

// ========================================
// GET HEADERS
// ========================================
function getHeaders_(sheet) {
const lastColumn =
sheet.getLastColumn();
if (lastColumn === 0) {
return [];
}
return sheet
.getRange(
1,
1,
1,
lastColumn
)
.getValues()[0]
.map(function(h) {


  return String(
    h || ""
  ).trim();

});


}

// ========================================
// SET VALUE
// ========================================
function setValue_(
row,
headers,
headerName,
value
) {
const index =
headers.indexOf(
headerName
);
if (index !== -1) {


row[index] = value;


}
}

// ========================================
// ROW → OBJECT
// ========================================
function rowToObject_(
headers,
row
) {
const obj = {};
headers.forEach(
function(header, i) {


  if (header) {

    obj[header] =
      row[i];

  }

}


);
return obj;
}

// ========================================
// PARSE POST DATA
// ========================================
function parsePostData_(e) {
if (!e) {
return {};
}
if (
e.postData &&
e.postData.contents
) {


const contents =
  e.postData.contents;


try {

  return JSON.parse(
    contents
  );

} catch (err) {

  // Try form data below

}


}
if (e.parameter && e.parameter.payload) {
  try {
    const obj = JSON.parse(String(e.parameter.payload));
    obj._fromForm = true;
    return obj;
  } catch (err) {}
}
if (e.parameter) {


return e.parameter;


}
return {};
}

// ========================================
// SECURITY HELPERS
// ========================================
function getAdminKey_() {
  return PropertiesService.getScriptProperties()
    .getProperty(ADMIN_KEY_PROPERTY) || "";
}

function isValidAdminToken_(token) {
  token = String(token || "").trim();
  if (!/^[0-9a-f-]{36}$/i.test(token)) return false;
  return CacheService.getScriptCache().get("adminToken:" + token) === "1";
}

function secureCompare_(a, b) {
  a = String(a || "");
  b = String(b || "");
  if (a.length !== b.length) return false;

  let result = 0;
  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

function candidateIdExists_(sheet, headers, candidateId) {
  const col = headers.indexOf("candidateId");
  if (col < 0 || sheet.getLastRow() < 2) return false;
  const ids = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues();
  return ids.some(function(r) {
    return String(r[0] || "").trim().toUpperCase() === candidateId;
  });
}

// ========================================
// GENERATE CANDIDATE ID
// ========================================
function generateCandidateId_(sheet, headers) {
  const col = headers.indexOf("candidateId");

  for (let attempt = 0; attempt < 20; attempt++) {
    const uuid = Utilities.getUuid().replace(/-/g, "").toUpperCase();
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    let randomPart = "";

    for (let i = 0; i < 6; i++) {
      randomPart += chars.charAt(parseInt(uuid.substr(i * 2, 2), 16) % chars.length);
    }

    const candidateId = "SP-26-" + randomPart.substr(0, 6);

    let exists = false;
    if (col >= 0 && sheet.getLastRow() >= 2) {
      const ids = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues();
      exists = ids.some(function(r) {
        return String(r[0] || "").trim() === candidateId;
      });
    }

    if (!exists) return candidateId;
  }

  throw new Error("Unable to generate candidate ID");
}

// ========================================
// SERVER-SIDE SCORING
// ========================================
// IMPORTANT:
// Replace ANSWER_KEY below with the authoritative question/answer mapping
// from your existing secure backend. Do NOT put it in the public HTML.
const ANSWER_KEY = {
  "During a society event, you notice that a member repeatedly makes comments toward another member that seem to be making them uncomfortable. The person hasn’t explicitly complained, but you notice them trying to avoid the member. What would you do?": {
    "Check in privately with the person who seems uncomfortable, listen without making assumptions, and, depending on their comfort, help them approach the appropriate society authority or reporting mechanism.": 40,
    "Speak privately to the person making the comments and let them know that their behaviour has been noticed and may be making someone uncomfortable, while encouraging them to be more mindful.": 30,
    "Keep an eye on the situation and speak to the person being affected if the behaviour continues or if they indicate that they want help.": 20,
    "Ask a few trusted society members whether they have noticed the same behaviour so you can better understand the situation before deciding how to proceed.": 10
  },
  "You have an important academic deadline tomorrow, while a society task you are responsible for is also due tomorrow. You realize you cannot give both tasks the attention they require. What would you do?": {
    "Prioritize based on urgency and consequences, communicate with your team as early as possible, and discuss whether the remaining work can be redistributed or the deadline adjusted.": 40,
    "Finish as much of the society task as possible and let the team know where things stand, while making sure your academic work is also completed.": 30,
    "Focus on your academic deadline first, then complete the society task afterward, even if it means submitting the society work later than expected.": 20,
    "Try to handle both tasks yourself by putting in extra hours, even if the quality of one or both tasks may suffer.": 10
  },
  "Your team is preparing for a major society event. One member has been doing most of the work, another has good ideas but rarely speaks up, and the rest of the team is struggling to agree on how to proceed. What would you do?": {
    "Organize a quick discussion where everyone can share their progress and concerns, actively involve the quieter member, and divide responsibilities according to people’s strengths and availability.": 40,
    "Encourage the member doing most of the work to delegate some responsibilities while giving the quieter member an opportunity to contribute their ideas.": 30,
    "Volunteer to take on some of the extra work yourself so the team can meet its deadlines, while allowing everyone else to continue with their existing responsibilities.": 20,
    "Suggest that everyone focus on their assigned tasks individually for now and revisit the disagreements once the event is closer.": 10
  },
  "Your first design decision is the wing. The aircraft needs to fly quickly and make rapid tight turns. You have three designs on the table:": {
    "A shorter, swept back wing with a relatively small area": 40,
    "a long, narrow, straight wing with a high aspect ratio": 30,
    "a very large, straight wing with a thick, highly cambered airfoil": 20,
    "a very short wing with almost no aerodynamic surface": 10
  },
  "You join the team and are given a task involving a part of RC aircraft design you’ve never worked with before. Nobody expects you to already know how to do it, but the team needs a first attempt within a week. What would you most likely do?": {
    "Find a few reliable resources, understand the basic principles, and make a small first attempt before asking for feedback.": 40,
    "Ask an experienced member to walk you through the entire process before starting.": 30,
    "Find a similar design online and adapt it until it seems suitable for our aircraft.": 20,
    "Ask to work on another part of the project where you already have experience.": 10
  },
  "the aircraft is almost complete. You know it will be performing aggressive high speed turns, so the entire aircraft body will experience substantial bending forces. You have enough carbon fibre reinforcement for only one location. Where would it be most valuable?": {
    "along the main wing spar": 40,
    "along the underside of the fuselage": 30,
    "for reinforcing the horizontal stabilizers": 20,
    "around the nose cone Phase 2: reverse engineering. Your aircraft performs successfully. During its first mission, you intercept one of the surveillance drones and manage to bring it down relatively intact. Back at the workshop, you place the captured aircraft beside your own. It is immediately obvious that its designers had very different priorities. Your job is to figure out what those priorities were so you can optimise your interceptors accordingly.": 10
  },
  "the captured drone has extremely long, narrow wings compared with its fuselage. Your friend argues that such large wings must have been chosen simply to generate more lift. Which explanation is more likely?": {
    "The wing’s high aspect ratio can reduce induced drag, making it more efficient during sustained flight": 40,
    "Long wings always produce less drag than short wings, regardless of speed": 30,
    "The designers wanted the drone to accelerate as quickly as possible": 20,
    "high aspect ratio wings prevent the aircraft from stalling": 10
  },
  "you discover that the drone uses a large diameter propeller turning relatively slowly. This seems strange for an aircraft designed to remain airborne for hours. What could be the reason?": {
    "It can move a large mass of air efficiently at relatively low flight speeds": 40,
    "Propeller diameter has little relationship with propulsion efficiency": 30,
    "Slow propellers require more energy, which improves endurance": 20,
    "Large propellers always produce greater top speed": 10
  },
  "While examining the drone, you notice that its ailerons are surprisingly small. Your friend concludes that the drone must therefore be poor at turning. Which is the better interpretation?": {
    "the aircraft may have been designed for gentle, efficient flight rather than rapid rolling manoeuvres": 40,
    "small ailerons prevent an aircraft from banking": 30,
    "The rudder becomes completely unnecessary when ailerons are small": 20,
    "Smaller ailerons always make an aircraft more manoeuvrable": 10
  },
  "Finally, you measure the drone’s wings, mass, cruise speed and battery capacity. You find that it has: low wing loading a highly efficient wing relatively low cruise speed a large battery modest control authority What does this combination most strongly suggest about the aircraft’s intended role?": {
    "Long-endurance surveillance": 40,
    "High-speed interception": 30,
    "Short-range aerobatics": 20,
    "High-altitude air racing": 10
  },
  "During the semester, the team is choosing between two designs for a mechanism. You think Design A is better, while another member strongly supports Design B. Both of you have reasonable arguments. The discussion is going nowhere. What would you do?": {
    "Agree on what the design actually needs to achieve and compare both options against those requirements.": 40,
    "Explain your reasoning again and try to convince them.": 30,
    "Ask the directors to choose so the team can move on.": 20,
    "Suggest building whichever design is easier to make.": 10
  },
  "During testing, you realize that a mistake you made earlier may have contributed to the aircraft behaving incorrectly. What do you do?": {
    "Tell the relevant team members what happened, help identify the effect, and work on the fix.": 40,
    "Tell a director privately and let them decide whether the rest of the team needs to know.": 30,
    "Wait until you’re certain that your mistake caused the problem before saying anything.": 20,
    "Try to fix it before mentioning it so you don’t unnecessarily slow everyone down.": 10
  },
  "You’re helping a high-school team during PSIFI. You’ve explained how to fix a particular issue twice, but they make the same mistake again. You have several other teams waiting for help. What do you do?": {
    "Explain the mistake again in a different way and have them make the correction themselves.": 40,
    "Give them the correct procedure to follow step-by-step and move on.": 30,
    "Fix it yourself so you can move on quickly.": 20,
    "Tell them to ask another volunteer because you have other teams to handle.": 10
  },
  "You finish your assigned part of the model earlier than expected. Another part of the aircraft is causing problems and the person responsible is struggling with it. What would you most likely do?": {
    "Ask the person what they’re stuck on and see whether you can help.": 40,
    "Tell a director that you’ve finished and ask for another assignment.": 30,
    "Start improving your own part further so it is as polished as possible.": 20,
    "Take over their part yourself so the project doesn’t fall behind.": 10
  },
  "Your robot works in simulation but fails in real life. Why?": {
    "Real world has things sim doesn't capture, like friction or delay": 40,
    "Robot is broken": 30,
    "Simulation was pointless": 20,
    "Real robots just don't work well": 10
  },
  "Robot only gets rewarded on full success, which is rare. What helps it learn faster?": {
    "Small rewards for getting closer to the goal": 40,
    "Letting it try randomly for a long time": 30,
    "Making the task easier": 20,
    "Skipping rewards, just showing the answer": 10
  },
  "What does \"sim2real\" mean?": {
    "Taking a model trained in simulation and using it on a real robot": 40,
    "Turning a real robot into a simulation": 30,
    "A type of robot hardware": 20,
    "Not sure": 10
  },
  "Robot learns by watching a human do a task. What's this called?": {
    "Imitation learning": 40,
    "Reinforcement learning": 30,
    "Supervised hardware training": 20,
    "Not sure": 10
  },
  "Teammate's code breaks your work before a deadline. What do you do?": {
    "Stay calm, find the change, fix it together": 40,
    "Undo their changes without telling them": 30,
    "Redo your own work from scratch": 20,
    "Wait for someone else to fix it": 10
  },
  "You find an ML concept you've never heard of. What do you do?": {
    "Look it up, understand the basics, ask questions if stuck": 40,
    "Skip it, hope it doesn't come up again": 30,
    "Wait for someone to explain later": 20,
    "Avoid the topic, assume it's too advanced": 10
  },
  "Why train a robot in simulation before real hardware?": {
    "Safer, faster, cheaper to fail and retry": 40,
    "Simulations are more fun": 30,
    "Real robots can't be trained": 20,
    "Not necessary, real-world training is always better": 10
  },
  "What excites you most about robots that learn?": {
    "Seeing something you built move in the real world": 40,
    "Coding in general": 30,
    "Looks good on a resume": 20,
    "Not sure yet, just curious": 10
  },
  "Prior experience with Python or coding?": {
    "Yes, written code before": 40,
    "A little, mostly tutorials": 30,
    "No, but willing to learn": 20,
    "No, and not sure I want to": 10
  },
  "You are assigned a task for the project, but a required material is unavailable and the team cannot buy new supplies. How do you move forward?": {
    "Look through leftover project supplies and scrap materials around the workspace to adapt or hand-build a working substitute.": 40,
    "Simplify the overall design plan so the project functions using only the materials already available on hand.": 30,
    "Pause physical work and focus entirely on planning and research until the required supplies become available.": 20,
    "Inform the team lead that your progress is blocked until the proper materials are provided.": 10
  },
  "You are responsible for completing a key deliverable due in two days, but unexpected complications mean you will not finish on time. What is your immediate reaction?": {
    "Inform the team immediately with an update on what is completed, where you are stuck, and a simplified version you can deliver by the deadline.": 40,
    "Work extra hours independently until the absolute last minute to get as close to completion as possible before letting anyone know.": 30,
    "Ask the team lead for a deadline extension so you can dedicate all your time to finishing the full deliverable properly.": 20,
    "Hand over your unfinished work to a teammate who has more free time so the team deadline isn't missed.": 10
  },
  "You finish your assigned project task two days ahead of schedule. A teammate working on a different part of the project is falling behind due to unexpected difficulties. What do you do?": {
    "Offer to take over their routine, non-technical workload (like documentation or organizing data) so they can focus entirely on solving their core problem.": 40,
    "Sit down with them to analyze their problem together and brainstorm potential workarounds.": 30,
    "Inform the team lead that you have completed your work early so you can be assigned your next individual task.": 20,
    "Use the extra time to refine and add extra polish to your own completed task beyond the initial requirements.": 10
  },
  "During a physical test, a custom component on the prototype snaps, and you do not have access to the workspace tools or machinery today to recreate it. How do you handle this setback?": {
    "Rig a temporary structural fix using simple items on hand (like zipties, tape, or scrap brackets) to keep non-structural testing moving forward today.": 40,
    "Pause physical testing on that section and update the design drawings on your computer to make the part stronger for when machinery is available.": 30,
    "Shift your focus entirely to writing project reports and documentation until you can access the workshop again.": 20,
    "Pack up the setup and wait until the workspace machinery reopens to resume work.": 10
  },
  "You and a teammate strongly disagree on how to organize the physical layout of the project components. Both ideas have valid merits, but arguing is stalling progress. How do you resolve it?": {
    "Agree on clear evaluation criteria (such as safety, ease of maintenance, and build time) to objectively choose the better option and move forward.": 40,
    "Build a quick, minimal test version of both ideas to see which one performs better in practice.": 30,
    "Escalate the decision to the team lead and agree to follow whichever option they choose.": 20,
    "Yield to your teammate’s layout preference so the team can maintain momentum and keep working without friction.": 10
  },
  "You are given a complex task that involves a discipline you have never worked with before. Your part of the work is due next week. How do you approach it?": {
    "Break the problem down into small manageable steps, research existing open-source examples, and build a simple basic working version first.": 40,
    "Reach out to a senior team member for a structured walkthrough before you begin working on a simplified attempt.": 30,
    "Search online for completed projects similar to yours and adapt their files directly to fit your team's requirements.": 20,
    "Ask the team lead to reassign you to a task that matches your existing skillset so you don't slow down the project.": 10
  },
  "While reviewing a teammate's setup before a test run, you notice a mistake that could potentially damage a key component. What is your immediate reaction?": {
    "Politely ask them to pause before starting, explain the potential risk, and work together for a few minutes to double-check the setup.": 40,
    "Let them run the test, but stay close by with safety precautions ready in case something goes wrong.": 30,
    "Step in immediately and fix the mistake yourself to ensure everything is completely safe before testing begins.": 20,
    "Notify a senior lead so they can inspect and approve the setup before any testing continues.": 10
  },
  "You are working late on a stubborn issue. The project isn't responding as expected, and after hours of trying, you are exhausted and unsure if the issue is a design flaw or human error. What do you do?": {
    "Document the exact conditions, steps taken, and errors in the team chat, then stop for the night to re-tackle it with a clear head in the morning.": 40,
    "Push through the fatigue and systematically test every variable one by one until you locate the issue before heading home.": 30,
    "Bypass the problematic section temporarily with a simplified workaround so overall project testing can continue overnight.": 20,
    "Replace the problematic component with a backup part from the workspace inventory to rule out part failure before calling it a night.": 10
  },
  "The team needs a custom mounting plate for a core device, but the main equipment needed to make it is booked by another team for the next two days. How do you keep the project moving?": {
    "Measure and manually cut a temporary baseplate from scrap wood or plastic using basic hand tools to allow initial testing to continue on the bench.": 40,
    "Temporarily secure the device using heavy-duty zipties and protective foam padding so functional testing isn't delayed.": 30,
    "Negotiate with the other team to trade time slots or share the equipment so you don't fall behind schedule.": 20,
    "Perform all non-physical planning and software tasks while waiting for the equipment to become free.": 10
  },
  "You are designing a competitive challenge for the event. A teammate pitches an elaborate puzzle that relies on obscure background lore that only hardcore fans would know. How do you approach this?": {
    "Frame the obscure lore as an optional \"bonus track\" or hint, keeping the core objective solvable through logical deduction.": 40,
    "Propose running a quick play-test with general participants to see if context clues make the puzzle decipherable without niche knowledge.": 30,
    "Recommend stripping out the obscure lore entirely and replacing it with a generic trivia question so casual teams don't feel lost.": 20,
    "Double down and tell the room that casual fans shouldn't bother showing up if they haven't memorized the background lore.": 10
  },
  "During a brainstorming session for an interactive mystery round, two teammates are locked in an argument over tone—one insists on dark psychological suspense, while the other wants campy humor. The discussion has stalled. What do you do?": {
    "Propose a unified premise that bridges both styles, starting with a lighthearted facade that takes a sudden suspenseful turn.": 40,
    "Redirect the focus by asking both members to walk through how their preferred tone specifically improves participant engagement.": 30,
    "Suggest splitting the round into two unrelated mini-modules so both members get their way, even if it bloats the schedule.": 20,
    "Delete the shared brainstorming doc in frustration and announce that you'll just write the entire script alone by midnight.": 10
  },
  "One hour before your auditorium event begins, a custom physical puzzle prop critical for the opening sequence breaks beyond immediate repair. What is your immediate course of action?": {
    "Convert the puzzle mechanic into a projected visual challenge on the main screens so the event starts on schedule.": 40,
    "Brief the room hosts on an interactive audience warm-up to buy time while the core puzzle is quickly adapted.": 30,
    "Try desperately to hot-glue the broken prop backstage, hoping it holds together for just the first five minutes.": 20,
    "Walk up to the microphone, announce that logistics ruined the round, and cancel the entire opening sequence on the spot.": 10
  },
  "During a live narrative round, a participant loudly points out that an official clue contradicts established franchise canon and claims the entire question is flawed. How do you handle this?": {
    "Stay composed, briefly verify the discrepancy, and make a transparent, uniform announcement to all teams so fairness is preserved.": 40,
    "Improvise in-character by framing the anomaly as a deliberate narrative twist, then note the issue for fair post-round scoring.": 30,
    "Politely tell the participant that the printed paper key is final and cannot be questioned until after the entire event concludes.": 20,
    "Publicly argue with the participant from the stage for five minutes to defend your writing credentials in front of the hall.": 10
  },
  "You notice a peer struggling to complete their clue-writing deliverables two days before the submission deadline due to sudden academic stress, though they have not posted in the group chat. What do you do?": {
    "Reach out to them privately to check in and offer to take two specific sections off their plate to ease the immediate pressure.": 40,
    "Suggest setting up a brief group co-working call so the sub-team can draft the remaining prompts together.": 30,
    "Message the department lead asking them to reassign the entire task to someone else without speaking to your peer first.": 20,
    "Post passive-aggressive screenshots of the untouched master sheet in the main group chat with a passive \"must be nice\" comment.": 10
  },
  "You have an idea for a cross-fandom tournament mechanic, but the senior organizing team seems comfortable repeating last year’s standard format. How do you present your case?": {
    "Create a concise one-page outline with a sample challenge and gameplay flow to show exactly how it works in practice.": 40,
    "Propose testing the concept as a low-stakes side quest or mini-round within the existing, proven framework.": 30,
    "Mention the idea briefly at the end of the meeting, but drop it completely the moment anyone raises a single doubt.": 20,
    "Secretly implement your mechanic into the master slides on event morning without telling the module heads.": 10
  },
  "Setup Week Commitment The week of the event requires daily on-ground setup. An unmovable academic commitment conflicts with two of your assigned setup shifts. How do you handle your availability?": {
    "Arrange a direct shift-swap with a teammate beforehand and commit to covering tear-down duty on the final evening.": 40,
    "Give the operations lead several days' advance notice along with a concrete list of tasks you will complete before and after your class.": 30,
    "Just drop a quick text in the group chat an hour before the shift saying you have class and asking someone to cover you.": 20,
    "Show up for five minutes to take an attendance photo, leave without telling anyone, and claim you were doing off-site errands.": 10
  },
  "Constructive Creative A module director reviews your first draft of narrative challenge prompts and notes that they feel \"dry and disconnected from the universe,\" requesting a revamp by tomorrow. How do you respond?": {
    "Ask for two specific examples of where the tone missed the mark, then revise the dialogue to reflect authentic character voices.": 40,
    "Partner with another writer on the team to bounce ideas around and infuse energetic world-building into the drafts.": 30,
    "Accept the critique quietly, but just add a few random catchphrases and character names into the existing draft without fixing the tone.": 20,
    "Send a defensive paragraph explaining that the director simply didn't understand the complex artistic vision behind the text.": 10
  },
  "An urgent request for three volunteers to finalize stage floor layouts tomorrow goes completely unanswered in the team chat for four hours. What do you do?": {
    "Drop a focused, upbeat message specifying the exact task: \"Need 2 people for 45 minutes at 4 PM to tape floor marks—who can jump in?\"": 40,
    "Tag a couple of teammates who mentioned being free around that time with a friendly, direct request to partner up.": 30,
    "Show up alone at 4 PM to do the whole layout yourself, secretly resenting everyone else for not volunteering.": 20,
    "Send a string of angry exclamation marks demanding to know why nobody respects the society's culture anymore.": 10
  },
  "When writing story prompts and challenge briefs for a multi-room fandom competition, what should your primary stylistic focus be?": {
    "Weaving in authentic in-universe dialogue and subtle details while keeping core objectives unmistakable for all participants.": 40,
    "Ensuring instructions are concise and foolproof first, using thematic terminology to color the narrative frame.": 30,
    "Writing long, elaborate paragraphs describing the room's atmosphere, even if it buries the actual rules of the puzzle.": 20,
    "Copy-pasting raw textbook Wikipedia summaries of plotlines with zero flavor, character dialogue, or creative engagement.": 10
  },
  "How is gradient descent used in ML?": {
    "It's used to visualize data distributions": 40,
    "It's an optimization algorithm used to minimize a model's loss function by iteratively adjusting parameters": 30,
    "It's used to split data into training and testing sets": 20,
    "It's a method for encoding categorical variables": 10
  },
  "How is linear regression different from logistic regression?": {
    "Linear regression is used for classification, logistic regression is used for prediction": 40,
    "There is no real difference; the names are interchangeable": 30,
    "Linear regression predicts continuous values, while logistic regression predicts probabilities for classification tasks": 20,
    "Logistic regression can only be used with one input variable": 10
  },
  "How many layers should a neural network have?": {
    "Always exactly 3 layers, no more and no less": 40,
    "There's no fixed number — it depends on the complexity of the problem, the data, and is often found through experimentation": 30,
    "Exactly 1 layer for all tasks": 20,
    "As many as possible, since more layers always improve performance": 10
  },
  "What is a word embedding?": {
    "A technique for removing punctuation from text": 40,
    "A numerical vector representation of words that captures semantic meaning and relationships between words": 30,
    "A method for translating text between languages": 20,
    "A way to count word frequency in a document": 10
  },
  "What type of content can be produced using machine learning models?": {
    "Only numerical predictions and classifications": 40,
    "Only images": 30,
    "A wide range of content — including text, images, audio, video, code, and structured data": 20,
    "Only text-based content like articles and summaries": 10
  },
  "An LLM gives a confident but factually incorrect answer. What is this commonly called?": {
    "Hallucination.": 40,
    "Overfitting.": 30,
    "Normalization.": 20,
    "Clustering.": 10
  },
  "You are participating in a 24-hour hackathon and want to build an AI-based application. What would generally be the best approach?": {
    "Start with a clear problem, build a small working prototype, test it, and then improve it based on time available.": 40,
    "Spend most of the hackathon training a large AI model from scratch regardless of the problem.": 30,
    "Build as many features as possible before testing whether the main idea works": 20,
    "Focus mainly on the presentation and leave the working prototype until the end.": 10
  },
  "Have you previously participated in a hackathon?": {
    "Yes, I have participated in one or more hackathons and have worked on a project as part of a team.": 40,
    "Yes, I have participated in a hackathon but had limited involvement in the technical/project development.": 30,
    "No, but I have worked on personal or academic projects involving programming/technology.": 20,
    "No, and I have not worked on any programming or technology-related project before.": 10
  },
  "You are given an unfamiliar AI/ML problem during a hackathon. What would be the best first step?": {
    "Understand the problem and requirements, research possible approaches, and then choose an appropriate solution based on available time and resources.": 40,
    "Immediately choose the most advanced model available.": 30,
    "Start coding immediately and figure out the problem while building the project.": 20,
    "Choose a model based only on how popular it is": 10
  },
  "The Wi-Fi goes down right as the main Tech Wars logic round begins. Participants at your assigned station are getting restless. What is your immediate response?": {
    "Announce to your section that there's a network issue and ask everyone to please sit tight and wait until it's resolved.": 40,
    "Wait in silence to see if the connection comes back on its own, avoiding eye contact so participants don't ask you questions you can't answer.": 30,
    "Pause the timer for your section, immediately text/inform the Event Head, and chat with your participants to keep them engaged while waiting for instructions.": 20,
    "Leave your assigned station immediately to hunt down campus IT so you can help get the network back up as fast as possible.": 10
  },
  "A highly competitive team disputes a penalty you gave them for a rule violation. They start arguing loudly with you in the middle of the hall.": {
    "Argue back firmly that your penalty is final and they need to sit down immediately or face disqualification.": 40,
    "Hand them a piece of paper, tell them to write down their exact dispute, and promise to hand it to the Event Head once the round finishes.": 30,
    "Quickly reverse the penalty yourself just to keep the peace and get them to stop yelling at you.": 20,
    "Politely ask them to lower their voices, stand your ground on the facts, but offer to call the Event Head over to review the dispute so the round isn't disrupted.": 10
  },
  "What type of games do you enjoy?": {
    "puzzle / platformer": 40,
    "Puzzle / strategy": 30,
    "action / adventure": 20,
    "casual / quickplay": 10
  },
  "We want to introduce a non-traditional \"wildcard\" round in Tech Wars this year. Which idea would you pitch to the team?": {
    "A tech-themed logic room with physical puzzles emphasizing teamwork and basic deduction over pure coding skills.": 40,
    "Letting the participants vote on the spot to see what kind of round they feel like doing.": 30,
    "A fast-paced tech trivia buzzer round covering famous CEOs and startup history.": 20,
    "A highly complex algorithmic challenge that only the most advanced programmers can solve.": 10
  },
  "You notice a quiet, introverted team member hasn't been doing much during event execution, while you are completely overwhelmed with tasks at your station.": {
    "Ignore them and struggle through the work yourself because explaining a task takes too much time.": 40,
    "Ask them if they want to shadow you and watch how you handle the rush so they know what to do next round.": 30,
    "Tell the Event Head that the member is free and can be assigned something.": 20,
    "Say, \"Hey, I have a lot of materials to organize for the next round right now. Could you help me sort them out?\"": 10
  },
  "The opening ceremony ran 45 minutes late, pushing the Tech Wars schedule back. A frustrated participant asks you what's going to happen to the remaining rounds. What do you do?": {
    "Tell them you don't know yet and that they should go find the Event Head to ask.": 40,
    "Tell them honestly that things are delayed, assure them the team is working on an updated schedule, and immediately check in with the Event Head for the official plan.": 30,
    ": Make a guess to settle the commotion despite it being true or not.": 20,
    "Take it upon yourself to announce to your entire section that the event will end 45 minutes late.": 10
  },
  "You suspect a participant is secretly using an AI tool to solve a puzzle during a strictly \"no internet\" round.": {
    "Call them out loudly in front of everyone to set a strict example for the other teams.": 40,
    "Interrupt them immediately, confront them about cheating, and demand to check their laptop.": 30,
    "Observe their screen discreetly for a few more minutes to be absolutely sure you have proof, then quietly flag down the Event Head to make the official call.": 20,
    "Stand visibly right behind their desk for the rest of the round so they get intimidated and close the tab themselves.": 10
  },
  "Day 1 of PSIFI is finally over. Everyone is exhausted, but the Tech Wars venue is a massive mess of wires, wrappers, and misplaced chairs.": {
    "Clean only your specific assigned desk area, grab your bag, and head out since you did your part.": 40,
    "Leave it for the campus janitorial staff, assuming they get paid to manage the halls anyway.": 30,
    "Go up to the Event Head and ask, \"What exactly do you want me to clean up first?\"": 20,
    "Pick up a trash bag, start clearing the area you are standing in, and casually ask the team members next to you if they want to help knock out this section together.": 10
  },
  "You discover a machine that can make predictions, but you have no idea how it works. You can give it examples and observe what it predicts. What would you enjoy doing most?": {
    "Give it different examples, look for patterns in its behaviour, and figure out what makes its predictions change.": 40,
    "If the machine already works, there's no reason to investigate how it got its answers.": 30,
    "Keep giving it examples until it starts producing the answers you want.": 20,
    "Try to discover the rules it seems to be following by comparing successful and unsuccessful predictions.": 10
  },
  "Which of the following set of CS domains intrigues you the most:": {
    "Cryptography, Cybersecurity & Computer Networks": 40,
    "Artificial Intelligence, Machine Learning & Computer Graphics": 30,
    "Artificial Intelligence, Cryptography & Game Development": 20,
    "Game Development, Computer Graphics & Computer Networks": 10
  },
  "You are testing a hydrogel and find that it absorbs much less water than expected. You have already spent several hours preparing the sample. What do you do next?": {
    "Check the preparation procedure and experimental conditions, identify possible sources of error, and repeat a controlled test before deciding that the material itself is unsuitable.": 40,
    "Repeat the same experiment with a fresh sample to determine whether the result was an isolated experimental error.": 30,
    "Modify the hydrogel formulation immediately and test whether the new version performs better.": 20,
    "Move on to another material because the current formulation does not appear to be performing well.": 10
  },
  "Your team has three potential hydrogel formulations, but only has enough chemicals and time to properly investigate one of them this week. How would you decide which one to prioritize?": {
    "Compare them using criteria relevant to the project's actual goal, such as water uptake, regeneration, durability, cost, and ease of preparation, then select the strongest overall candidate.": 40,
    "Choose the formulation that appears most promising based on the available research and test it thoroughly.": 30,
    "Choose the easiest formulation to prepare so the team can obtain experimental results quickly.": 20,
    "Choose the formulation with the highest reported water uptake because maximizing water absorption is the most important objective.": 10
  },
  "You have collected experimental data for water uptake, but one measurement is significantly different from the others. What do you do?": {
    "Check the experimental conditions and measurement process, then repeat the relevant trial before deciding whether the unusual value should influence the conclusions.": 40,
    "Repeat the experiment several times and compare the new results with the original data.": 30,
    "Keep the value in the dataset but mention that it appears unusual when presenting the results.": 20,
    "Remove the value because it does not agree with the other measurements.": 10
  },
  "Your hydrogel performs extremely well under the conditions in which you tested it, but you realize those conditions are very different from the climate in which AQUA is ultimately intended to operate. What would you do?": {
    "Identify the environmental factors that could affect performance and design experiments that test the material under more realistic conditions.": 40,
    "Repeat the existing experiments several times first to make sure the original results are reliable.": 30,
    "Research the expected environmental conditions and use published data to estimate how the hydrogel might perform.": 20,
    "Continue using the current conditions because they allow the team to compare different materials consistently.": 10
  },
  "Your team has developed a hydrogel that works well, but it becomes difficult to handle once it has absorbed water. The next stage is to design a device around it. What would you prioritize?": {
    "Identify the practical requirements the hydrogel holder must satisfy—such as support, airflow, drainage, water collection, and ease of replacement—and design around those constraints.": 40,
    "Create a simple holder first and use prototype testing to identify what needs to be improved.": 30,
    "Look at existing water-harvesting devices for design ideas and adapt features that appear useful.": 20,
    "Design the most compact holder possible so that the final device takes up minimal space.": 10
  },
  "Your first device prototype collects water, but the amount is significantly lower than what the hydrogel produced during laboratory testing. What would you investigate?": {
    "Break the system into stages and determine where the loss is occurring—for example, moisture reaching the hydrogel, water release, drainage, or collection—before changing the design.": 40,
    "Compare the conditions inside the device with those used during the hydrogel's laboratory testing.": 30,
    "Modify the physical arrangement of the hydrogel inside the device and see whether water collection improves.": 20,
    "Increase the amount of hydrogel in the device so that more water can potentially be collected.": 10
  },
  "You and another team member propose two different designs for the hydrogel-holding section of AQUA. Both designs are plausible, but you cannot determine which will perform better just by looking at them. What do you do?": {
    "Define the key performance criteria and create a simple test or prototype that allows the two designs to be compared objectively.": 40,
    "Discuss the advantages and disadvantages of both designs and agree on the one with the stronger overall argument.": 30,
    "Ask the project lead to choose which design the team should pursue.": 20,
    "Combine both designs into one so that neither idea is rejected.": 10
  },
  "You are asked to help optimize the device, but you have never used simulation or CFD software before. The rest of the team is already working on other parts of the project. What do you do?": {
    "Learn the fundamentals needed for the specific problem, study simple examples, and build up from a basic model before attempting the complete device.": 40,
    "Ask someone experienced with CFD to explain the basic workflow and then attempt a simplified version yourself.": 30,
    "Find an existing simulation similar to the project and try to modify it for AQUA.": 20,
    "Ask to be moved to an experimental task because CFD is outside your current skillset.": 10
  },
  "Your CFD simulation suggests that a particular device geometry should perform better, but when you build and test it physically, the opposite happens. What would you do?": {
    "Compare the assumptions and boundary conditions in the simulation with the real experimental conditions and use the discrepancy to improve the model.": 40,
    "Repeat the physical experiment to determine whether the experimental result is reproducible.": 30,
    "Run simulations for several other geometries to see whether the trend remains consistent.": 20,
    "Assume the physical prototype is inaccurate because the simulation predicted the better-performing design.": 10
  },
  "You are working on the prototype and discover that a small design change could improve performance, but it would require rebuilding a component that took your team several days to make. What do you do?": {
    "Estimate the potential benefit and effort involved, then test the idea with a simplified version or analysis before deciding whether a full rebuild is justified.": 40,
    "Discuss the potential improvement with the team and decide whether it is significant enough to justify rebuilding the component.": 30,
    "Leave the current prototype unchanged because rebuilding it could delay the project.": 20,
    "Immediately rebuild the component because improving performance should take priority over the existing work.": 10
  },
  "Your team has spent two weeks developing a particular approach, but new results suggest that another approach may ultimately work better. Switching would mean losing some of the work already completed. What would you do?": {
    "Compare both approaches against the project's objectives and use the evidence to decide whether switching is worth the time and resources already invested.": 40,
    "Run a small test of the alternative approach before deciding whether to abandon the current one.": 30,
    "Continue with the current approach for now because changing direction this late could create unnecessary delays.": 20,
    "Switch immediately because the newer approach appears more promising.": 10
  },
  "You finish your assigned hydrogel experiments earlier than expected while another member is struggling to process their experimental results. What do you do?": {
    "Ask what part of their workload is slowing them down and offer specific help while making sure your own results and documentation are complete.": 40,
    "Sit with them and help troubleshoot the problem so they can continue independently.": 30,
    "Inform the project lead that your work is complete and ask whether there is another task you can take on.": 20,
    "Use the remaining time to further improve your own experiment even if the additional work is not currently needed.": 10
  },
  "You are responsible for recording experimental results, but you notice that the team has been using slightly different procedures between trials. What do you do?": {
    "Raise the issue, identify which experimental conditions need to be standardized, and help establish a consistent procedure for future trials.": 40,
    "Document the differences carefully so that they can be accounted for when interpreting the results.": 30,
    "Continue collecting data but make a note that the experimental procedure has varied.": 20,
    "Use the results as they are because some variation is unavoidable in experiments.": 10
  },
  "You are given a task for AQUA that seems straightforward, but you realize halfway through that you don't understand why the task matters to the overall project. What do you do?": {
    "Understand how the task connects to the project's larger objective before continuing, and ask targeted questions if necessary.": 40,
    "Complete the task first and then ask how the result will be used in the project.": 30,
    "Follow the instructions exactly because understanding the broader purpose is not necessary to complete your assigned work.": 20,
    "Ask to be assigned a different task whose purpose is clearer to you.": 10
  },
  "eventually produces a prototype that works, but it is too expensive and complicated to realistically deploy in the communities the project is targeting. What would you consider the most important next step?": {
    "Identify which aspects of the design drive cost and complexity and redesign around the actual constraints of the intended users.": 40,
    "Investigate cheaper materials or manufacturing methods that could preserve most of the prototype's performance.": 30,
    "Improve the prototype's performance further first, then work on reducing its cost.": 20,
    "Consider the prototype successful because demonstrating that the technology works is the main objective of the project.": 10
  },
  "On event day, you notice we are running 1 hr behind schedule, what do you do?": {
    "😎 Take your time to figure out what’s causing the delay and mitigate the problem": 40,
    "⚡️⌛️ Change the schedule for the upcoming days to cater for the delays": 30,
    "😏😇 Just try and rush every process, telling participants to hurry up": 20,
    "Be nonchalant, and act like it’s not your problem": 10
  },
  "Halfway through a group task, it’s clear one teammate hasn’t touched their part, and the deadline’s two days out. What’s your call?": {
    "Talk to them directly to understand what’s going on": 40,
    "Quietly take on the extra work yourself to stay on track": 30,
    "Let the team lead know and leave it to them": 20,
    "Give it more time in case they’re just running behind": 10
  },
  "A participant comes to you furious because they believe the referee made an unfair call. 🧏🧏‍♀️ What do you do?": {
    "🤬🤬 Hear them out, check the relevant rule, and involve the Event Heads": 40,
    "🤷 Argue about the interpretation of the rule": 30,
    "😡😡 Tell them the ref’s decision is final and walk away": 20,
    "Fist fight": 10
  },
  "You’re helping build something mechanical for an event, and a part you assembled clearly isn’t holding together the way it’s supposed to. You’re not totally sure why. What’s your move?": {
    "Take it apart, find the actual weak point, and test a proper fix": 40,
    "Add some tape or glue to hold it for now and fix it later": 30,
    "Ask someone more experienced to take a look and fix it": 20,
    "Leave it alone since it’s still technically working fine": 10
  },
  "Partway into a task, you realize a different approach would work much better — but it means redoing finished work. What do you do?": {
    "Bring the better idea to your lead, even if it means backtracking": 40,
    "Finish the current approach, then suggest the idea for next time": 30,
    "Just switch over to the new approach on your own": 20,
    "Stick with the original plan since it’s already approved": 10
  },
  "Two days before the big showcase, you spot a flaw in the arena setup that could make a round unfair for some teams. Fixing it eats into most of your remaining time. What now?": {
    "Flag it to your lead now and push to fix it properly": 40,
    "Quietly fix it yourself overnight without telling anyone": 30,
    "Leave it, since changes this late could cause more issues": 20,
    "Wait to see if anyone else notices it first": 10
  },
  "Partway through a project, costs are running higher than planned and the budget won’t stretch to everything. What do you do?": {
    "Sort essentials from nice-to-haves and rebuild the plan now": 40,
    "Ask if there’s room to increase the budget first": 30,
    "Trim a little from every item so nothing’s fully cut": 20,
    "Keep going as planned and deal with it later": 10
  },
  "An electronic component isn’t responding the way it should during testing, and there’s no clear error message. How do you approach figuring out what’s wrong?": {
    "Check the simplest causes first, like power and connections": 40,
    "Search the exact symptom online and try the first fix you find": 30,
    "Swap the part out and hope the new one works fine": 20,
    "Ask a teammate to take a look since you’re unsure": 10
  },
  "It’s mid-semester, Robowars is kicking into gear, but so are your academics. You 😭 volunteered for a task from the Event Head because you thought you were Goated. Now 😭🥲 you’re overwhelmed, with 3 quizzes and one assignment due. What do you do?": {
    "🫡 🐐 Ask your Head for help and an extension, explaining your academic load": 40,
    "😝😍 Decide you’re invincible and just do both, leaving the same amount of work": 30,
    "🗣️🗣️🗣️ Quickly skim content for your quizzes and rush the Robowars task": 20,
    "Sacrifice your life for Robowars — screw the quizzes": 10
  },
  "Something breaks in front of a crowd during a live event, through nobody’s fault. What’s your first instinct?": {
    "Stay calm, explain what’s happening, and start troubleshooting": 40,
    "Try to fix it quietly without drawing more attention": 30,
    "Step back and let a senior teammate take over instead": 20,
    "Give it a few minutes to see if it resolves itself": 10
  },
  "Your lead hands you a task with almost no detail — just “make it exciting.” Where do you start?": {
    "Ask a couple of pointed questions to pin down the goal": 40,
    "Build a rough version and check in early for feedback": 30,
    "Look at how similar things were done before and follow that": 20,
    "Hold off until you’re given clearer instructions": 10
  },
  "Three tasks land on your plate at once for different parts of the event, and there’s no way to do all three justice. How do you decide?": {
    "Weigh which task matters most and flag the tradeoff early": 40,
    "Split your effort evenly so nothing is fully dropped": 30,
    "Go by whichever deadline happens to hit first": 20,
    "Stick to whichever one you’re most confident about": 10
  },
  "An hour before an event, the main microphone setup is faulty and the vendor refuses an immediate replacement, claiming it worked on delivery. What is your immediate course of action?": {
    "Reallocate a secondary microphone from the backstage briefing room to start on time, while sending a team member to pick up a replacement from a nearby vendor using petty cash.": 40,
    "Utilize the venue’s house audio system as an immediate fail-safe, while documenting the faulty vendor gear to claim a full refund post-event.": 30,
    "Contact the vendor's senior account manager to demand an emergency replacement be dispatched immediately before the doors open.": 20,
    "Instruct the event hosts to proceed with lapel mics only and adjust the main soundboard frequencies to boost the audio output.": 10
  },
  "Thirty minutes before an event, another department demands 50 extra chairs and a projector that were not in their original requisition form. How do you respond?": {
    "Release 25 spare chairs from storage and a backup projector, while explaining that additional items cannot be supplied without compromising other venue setups.": 40,
    "Fulfill the request by pulling unused seating from the registration overflow zone and borrowing a projector from the media team.": 30,
    "Advise them to adjust their room layout to standing-room style for the extra guests, as unscheduled inventory cannot be deployed right before start time.": 20,
    "Direct their logistics liaison to locate and transfer extra furniture from the secondary holding hall independently.": 10
  },
  "Right before a deadline, a team member publicly refuses to follow your setup instructions and creates tension within the team. How do you manage this?": {
    "Brief them privately on the operational priority, reassign them to oversee the stage backdrop alignment independently, and review the behavioral issue after event teardown.": 40,
    "Pause work for two minutes to listen to their suggested alternative setup, adopt valid points to build consensus, and resume execution.": 30,
    "Step in to finish their technical task yourself and assign them to monitor incoming attendee queues at the main entrance instead.": 20,
    "Remind them firmly of the deadline in front of the team, insist on compliance with the original plan, and evaluate their position post-event.": 10
  },
  "A lead from another department behaves entitled, orders your team around, and demands priority over the master schedule. What do you do?": {
    "Intercept the lead personally, take over communication regarding their requests, and instruct your team to stick strictly to the master timeline.": 40,
    "Fulfill their setup adjustment immediately to maintain inter-departmental harmony, then adjust your remaining timeline accordingly.": 30,
    "Escalating the situation to their department head immediately so they can instruct their lead to follow protocol.": 20,
    "Tell your team members to politely inform the lead that all requests must be submitted through official written channels first.": 10
  },
  "Two simultaneous events request your society's single backup generator. How do you resolve who gets it?": {
    "Direct the generator to the main stage based on its higher total wattage requirement, while rerouting secondary venue power from the main grid for the second event.": 40,
    "Award the generator to the department that submitted their technical requisition form first according to official timestamp records.": 30,
    "Split the backup generator's total capacity between essential sound systems at both venues using heavy-duty splitter cables.": 20,
    "Request both department heads to meet immediately and decide between themselves who requires the backup unit more.": 10
  },
  "Two days before a 500+ person event, your printing vendor cancels all delegate cards and certificates. What is your primary strategy?": {
    "Split the order between two local print shops for immediate turnaround, while prioritizing delegate passes for day one and delivering certificates digitally post-event if needed.": 40,
    "Negotiate a 30% rush surcharge with a commercial print firm to clear their queue and process the full 500-unit order within 36 hours.": 30,
    "Purchase pre-cut cardstock to print delegate passes using campus high-speed printers, while outsourcing only the certificates to an external vendor.": 20,
    "Utilize an online emergency printing service with express overnight courier shipping to deliver all materials directly to the venue.": 10
  },
  "Mid-setup, you discover a vendor delivered 30% fewer setup materials than paid for, and you have no remaining budget. How do you proceed?": {
    "Contact the vendor with the signed delivery receipt to demand emergency dispatch, while re-spacing existing stage framing and backdrops to maintain visual coverage safely.": 40,
    "Pause setup on the main stage, consolidate all delivered materials toward the entry display, and wait for the vendor to send the remainder.": 30,
    "Reallocate funds from the post-event refreshments budget to purchase the missing materials directly from a local hardware store.": 20,
    "Proceed with the original assembly plan, leaving minor structural gaps in less visible areas behind the backdrop.": 10
  },
  "High-value equipment goes missing from an unlocked storage room during setup due to a team member's oversight. What are your immediate actions?": {
    "Source temporary replacement units from secondary department kits, lock the storage facility under a strict sign-in log, and conduct an internal inventory audit after the event.": 40,
    "Pause non-essential setup duties for 15 minutes to conduct a sweep of the immediate venue, then assign a dedicated guard to the storage door.": 30,
    "Report the missing items to venue security immediately, reassign the responsible team member off-site, and source spare equipment using emergency funds.": 20,
    "Lock the storage room immediately, restrict key access to yourself, and require the responsible team member to track down the missing items.": 10
  },
  "Your team of 5 logistics members must cover a grueling 12-hour event shift (setup, active event, and teardown). How do you approach creating the shift schedule and assigning responsibilities fairly?": {
    "Draft a tiered schedule with rotating core roles, present it to the team to gather preferences for peak operational hours, and resolve conflicts based on individual strengths and availability.": 40,
    "Divide the 12 hours into equal, fixed time blocks and assign team members randomly to ensure total impartiality.": 30,
    "Allow team members to self-select their preferred shifts on a first-come, first-served basis, stepping in only if a shift remains unstaffed.": 20,
    "Assign the hardest tasks (setup and teardown) to the most experienced members and let newer members handle the active event shifts.": 10
  },
  "During a fast-paced event, you notice two of your team members are completely overwhelmed managing stage logistics, while two others assigned to information desk duties have very little crowd to handle. How do you rebalance the workload?": {
    "Reassign one information desk member to assist with stage setup tasks immediately, while briefing both teams on cross-departmental coverage so operations stay fluid.": 40,
    "Step in personally to handle the stage logistics duties yourself so you don't disrupt either team's original assignments.": 30,
    "Instruct the two stage logistics members to take a short break while you rotate the information desk members into their roles entirely.": 20,
    "Tell the stage logistics team to hang in there until their scheduled shift change, as shifting roles mid-event causes confusion.": 10
  },
  "It's the night of a big stargazing event and the sky suddenly clouds over 30 minutes before it starts. What's your first move?": {
    "Start looking up indoor backup activities to keep the crowd engaged": 40,
    "Announce a possible delay to attendees and gauge their patience level": 30,
    "Check weather updates repeatedly hoping it clears up": 20,
    "Ask a senior member what the plan is and wait for instructions ---": 10
  },
  "A visitor asks you a question about space you don't know the answer to. You:": {
    "Admit you don't know and offer to find out and get back to them": 40,
    "Use the moment to spark a discussion and explore the answer together": 30,
    "Redirect them to a senior member confidently": 20,
    "Give your best guess so they don't feel unanswered ---": 10
  },
  "You're assigned to explain how a telescope works to a group of children. How do you prepare?": {
    "Prepare a short demo/activity so they learn by doing": 40,
    "Think of simple analogies from everyday life to explain the concept": 30,
    "Keep it minimal and let questions guide the explanation": 20,
    "Memorize technical facts and use scientific terms to sound credible ---": 10
  },
  "During a team meeting, two members disagree on how an event should be run. You usually:": {
    "Try to understand both sides before sharing your opinion": 40,
    "Suggest testing both ideas on a smaller scale first": 30,
    "Push for a quick decision so planning isn't delayed": 20,
    "Stay quiet and go with whatever the majority decides ---": 10
  },
  "Which of these excites you most about working in an observatory team?": {
    "Getting to explain complicated ideas in a way anyone can understand": 40,
    "Being hands-on with telescopes and technical setups": 30,
    "The unpredictability — every night sky and every crowd is different": 20,
    "Organizing and executing large-scale public events ---": 10
  },
  "A stargazing session has way more attendees than expected. Your first priority is:": {
    "Coordinating quickly with your team to divide responsibilities": 40,
    "Crowd management so everyone gets a fair turn": 30,
    "Improvising additional activities to keep everyone engaged": 20,
    "Making sure the core content/experience isn't compromised ---": 10
  },
  "How do you usually handle tasks with tight, fixed deadlines (like event day prep)?": {
    "Prioritize the most critical tasks and adapt as needed": 40,
    "Make a checklist and work through it systematically": 30,
    "Delegate and rely on teamwork to divide the load": 20,
    "Work in bursts of high focus closer to the deadline ---": 10
  },
  "A visitor is skeptical and asks \"why does this even matter/how is this useful?\" How do you respond?": {
    "Explain the real-world applications and importance of astronomy": 40,
    "Acknowledge the question and give a short, honest perspective without over-selling it": 30,
    "Share why you personally find it fascinating": 20,
    "Ask them why they came if they aren't interested ---": 10
  },
  "When learning to operate new equipment (like a telescope) for the first time, you're most likely to:": {
    "Ask someone experienced to show you hands-on": 40,
    "Read the manual/instructions thoroughly first": 30,
    "Watch others do it a few times before attempting it": 20,
    "Try it yourself and learn through trial and error ---": 10
  },
  "It's a clear night and the event is going smoothly, but you notice one attendee seems bored/disengaged. You:": {
    "Approach them and try to get them more involved": 40,
    "Adjust your own explanation/activity to make it more interesting for everyone": 30,
    "Mention it to a teammate to help re-engage them together": 20,
    "Let them be — not everyone needs to be equally engaged ---": 10
  },
  "How would you rate your comfort with public speaking/explaining things to strangers?": {
    "Very comfortable, I enjoy it": 40,
    "Comfortable once I know the topic well": 30,
    "A bit nervous, but I manage": 20,
    "I'd rather support behind the scenes ---": 10
  },
  "A last-minute change is announced — the event location/time shifts an hour before it starts. Your reaction is to:": {
    "Immediately start informing/coordinating with attendees and team": 40,
    "Suggest a way to make the transition smoother for everyone": 30,
    "Take a moment to understand the reason before reacting": 20,
    "Focus on your specific responsibility and trust others to handle the rest ---": 10
  },
  "If given creative freedom to design one segment of a stargazing event, what would you focus on?": {
    "An interactive activity for the audience": 40,
    "A compelling story/narrative around the night's sky": 30,
    "Visual aids or props to make it more engaging": 20,
    "A simple, well-explained core experience without added frills ---": 10
  },
  "You've been assigned a repetitive task (like data entry for event registrations) alongside a more \"exciting\" one. You:": {
    "Use AI to quickly finish the repetitive one, and move on to the exciting one": 40,
    "Do the repetitive task first to get it out of the way": 30,
    "Ask if the repetitive task can be handled differently/faster": 20,
    "Do the exciting task first since it needs more mental energy anyway ---": 10
  },
  "An event you helped plan doesn't go as expected (low turnout/technical issues). Your first thought is:": {
    "How do we salvage the current situation": 40,
    "What can we learn from this for next time": 30,
    "Whether the planning process itself needs to change": 20,
    "Reassuring the team that it's okay, these things happen ---": 10
  },
  "What matters most to you when working as part of a team on a shared goal?": {
    "Clear communication and coordination": 40,
    "Everyone contributing fairly": 30,
    "Getting the outcome right, regardless of process": 20,
    "Enjoying the process and staying motivated together.": 10
  },
  "Classroom management / patience You're mid-workshop teaching a group of 10-year-olds a science experiment. One kid keeps interrupting, distracting others, and clearly isn't engaged with the activity.": {
    "Give them a specific task within the experiment (e.g. “you’re in charge of timing this step”) to redirect their energy": 40,
    "Have a quiet one-on-one word with them about focusing, without singling them out in front of the group": 30,
    "Ask a fellow volunteer to work with them individually so the main session isn’t slowed down": 20,
    "Let it go for now — some kids just need more time to warm up, and stopping to address it disrupts the rest of the group": 10
  },
  "Adaptability under pressure You've spent two weeks planning a workshop with a partner NGO. The day before, they tell you the venue fell through and you now have half the space and half the time you planned for.": {
    "Redesign on the spot into a different format (e.g. smaller rotating groups) that fits the new constraints": 40,
    "Cut the workshop down to the highest-impact activity and drop the rest": 30,
    "Keep the full plan but compress each activity, even if it means rushing some parts": 20,
    "Push to reschedule so the workshop can be delivered as originally designed": 10
  },
  "Resourcefulness / creativity You're assigned to teach a STEM concept to a group of kids, but you arrive and realize half the materials you were promised weren't provided.": {
    "Improvise using whatever’s around you (classroom objects, phones, anything) to substitute for the missing materials": 40,
    "Simplify the experiment so it works with only what’s on hand, even if it’s less impressive": 30,
    "Shift to a more discussion/demonstration-based version of the lesson instead of hands-on": 20,
    "Do the parts you can with the materials you have, and clearly explain to the kids what the rest would have looked like": 10
  },
  "Teamwork / conflict resolution You're part of a 4-person team preparing a workshop. One teammate has consistently missed prep meetings and hasn't done their assigned part, and the event is in 3 days.": {
    "Talk to them directly and ask what’s going on before assuming anything": 40,
    "Give them one more chance with a firm, specific deadline before redistributing the work": 30,
    "Flag it to your team lead so it’s handled at a level above just the team": 20,
    "Redistribute their part among the rest of the team quietly and deal with the conversation after the event": 10
  },
  "Empathy / mentorship While mentoring a student for something like PsiFi, they tell you they don't think they're \"smart enough\" to be good at STEM and seem ready to give up.": {
    "Ask them what part specifically feels hardest, so the conversation is about the real obstacle instead of a general feeling": 40,
    "Point out specific things they’ve already done well to show them evidence against that belief": 30,
    "Shift focus away from ability entirely and onto effort — that improvement, not talent, is what actually matters here": 20,
    "Share that struggling is normal and even people who seem naturally good at it started out confused too": 10
  },
  "Prioritization / time management you have one week and a limited budget to plan a workshop. You can either spend most of your effort making the content really strong, or spend it on logistics (transport, permissions, NGO coordination) to make sure it actually happens smoothly.": {
    "Prioritize logistics — a great lesson plan is worthless if the workshop falls apart before it starts": 40,
    "Split effort evenly across both, accepting that neither will be as polished as it could be": 30,
    "Do a simpler, lower-effort version of the content so you can fully focus on logistics without risk": 20,
    "Prioritize content — a workshop that runs a bit rough but teaches something meaningful is still a win": 10
  },
  "Initiative / judging fit for non-technical contributors You're a non-STEM student who wants to join SRP. During onboarding, you're told the team currently needs help most with lesson planning and teaching delivery, not building the technical experiments themselves.": {
    "Volunteer for teaching delivery since that’s clearly where the need is, even if it feels outside your comfort zone": 40,
    "Offer to help simplify and explain the technical content in plain language, treating it as a translation problem": 30,
    "Ask if there’s a hybrid role — helping shape how the experiment is presented even without designing it": 20,
    "Take it as a chance to learn the technical side too, so you can eventually contribute there as well": 10
  },
  "Reflective judgment / defining success A workshop you ran went exactly according to plan: every activity finished on time, nothing went wrong logistically. But afterward, most kids seemed bored and weren't really engaging with the material.": {
    "Reserve judgment until getting feedback directly from the kids or the NGO partner about how it landed": 40,
    "Look for the exact moment engagement dropped and treat it as a specific fixable design flaw, not a verdict on the whole session": 30,
    "Call it a partial success — smooth execution matters, and engagement can be improved next time without it being a failure": 20,
    "Call it a failure — if the kids weren’t engaged, the actual goal wasn’t met regardless of how smoothly it ran": 10
  },
  "A team complains that their assigned prompt is too confusing and asks you to just tell them what to do. How do you help?": {
    "Ask them a series of guiding questions about the prompt's keywords to help them arrive at their own creative interpretation without giving them the answer.": 40,
    "Break down the prompt into simpler terms and give them one broad, hypothetical example to kickstart their brainstorming process.": 30,
    "Tell them that interpreting the complex abstraction is the entire point of the challenge and walk away to maintain strict fairness.": 20,
    "Allow them to swap their prompt for an easier one with a point penalty, complicating the grading process for the judges.": 10
  },
  "Three deadlines land in the same week: a deliverable for Pixel, a quiz, and a group project. What actually happens?": {
    "You knock out the quick wins to clear mental space, then focus on the big one.": 40,
    "You do the hardest thing first while your energy is highest.": 30,
    "You map out the week hour by hour and stick to it.": 20,
    "You tell someone early that one of the three will be late, and pick which.": 10
  },
  "You receive a submission that is visually breathtaking but completely fails to incorporate the required scientific concept. How should it be judged?": {
    "Score the design elements highly, but apply a strict penalty for missing the core scientific prompt to maintain fairness to other teams.": 40,
    "Use the given explanation while submitting let the team express correctly if the science is hidden in a subtle visual metaphor, failing that specific criteria only if they cannot defend it.": 30,
    "Fail the submission entirely; no matter how good it looks, ignoring the fundamental brief makes the entry invalid.": 20,
    "Overlook the missing science because the stunning visuals elevate the overall quality of the event's showcase.": 10
  },
  "A team claims another group overheard their brainstorming and stole their central idea for the design challenge. How do you handle the accusation?": {
    "Neutrally explain that parallel thinking is common under tight constraints and, without hard proof, neither team will be penalized, shifting their focus back to out-executing the other team.": 40,
    "Quietly notify the Event Head and the judges to scrutinize both final submissions for originality, without interrupting either team's workflow during the round.": 30,
    "Separate both teams and ask them to independently explain the origin of their concepts to the Event Head, which unnecessarily halts their progress and escalates the drama.": 20,
    "Immediately confront the accused team and demand to see their working files to prove they didn't cheat.": 10
  },
  "You need to explain a creative concept to someone who doesn't work in visuals at all. How do you do it?": {
    "Show them a reference from something they already know.": 40,
    "Sketch it out roughly on the spot.": 30,
    "Explain the feeling you want people to have, not the design itself.": 20,
    "Skip the explanation and just build a quick mockup.": 10
  },
  "Your event head is visibly stressed, frantically trying to fix a scheduling conflict, and snapping at people. How do you support them?": {
    "Step in and silently take over their minor logistical tasks so they have the mental space to solve the main crisis.": 40,
    "Wait for a brief pause, then calmly offer to take the scheduling conflict off their plate entirely so they can manage the rest of the room.": 30,
    "Give them space and stay completely out of their way, leaving them to handle the stress alone so you don't risk making it worse.": 20,
    "Rally the rest of the organizing team to give the head a pep talk, wasting time and drawing unnecessary attention to their panic.": 10
  },
  "During a brainstorming meeting, two loud, confident members are dominating the conversation. Three quieter members haven't spoken at all. The ideas being discussed are just \"okay.\"": {
    "Pivot the meeting to a \"silent brainstorm,\" having everyone write their ideas on sticky notes for 5 minutes to neutralize the loud voices and equalize input.": 40,
    "Politely interrupt the loud members and explicitly ask the quieter members for their feedback on the ideas just proposed, bringing them directly into the fold.": 30,
    "Wait until after the meeting and take the quiet members aside privately to ask for their ideas, though this fails to fix the team dynamic at the moment.": 20,
    "Let the loud members run the whole meeting because their enthusiasm will drive the execution, even if the ideas remain mediocre.": 10
  },
  "A teammate stops responding two days before submission. Their part is unfinished. What's your move?": {
    "Keep trying to reach them while starting a stripped-down backup version.": 40,
    "Redistribute their work across the rest of the team and move on.": 30,
    "Finish their part yourself and sort it out afterwards.": 20,
    "Escalate to the department head immediately.": 10
  },
  "A participating team loses all their digital files due to a software crash 30 minutes before the deadline. How do you respond?": {
    "Calm the team down and instruct them to rapidly sketch their concepts on paper and write a brief text summary so they still have a structured, physical submission for the judges to evaluate.": 40,
    "Enforce the strict deadline, but allow them to submit whatever rough notes or partially saved assets they can recover, evaluating them solely on what is turned in.": 30,
    "Give them a quick 15-minute extension to rebuild what they can, even though it delays the master schedule and gives them a slight time advantage over others.": 20,
    "Pause the entire competition for all teams to let them recover their data out of fairness to their hard work.": 10
  },
  "You are reviewing a physics round intended for high-school students. During internal test-solving, you find that one question is consistently taking significantly longer than expected, while another is being solved almost immediately by everyone. The problem setter argues that the difficult question will help separate the strongest teams. What is the best approach?": {
    "Keep both questions, since having a few questions that strongly distinguish the top teams is useful in a competitive round.": 40,
    "Replace the difficult question with a moderately challenging one and keep the easier question, ensuring that participants are able to progress through the round without getting stuck.": 30,
    "Have additional team members independently test both questions under realistic time constraints, then use their performance to decide whether either question needs modification before finalizing the paper.": 20,
    "Evaluate both questions in the context of the round's intended difficulty progression, marks, expected solving time, and overall balance; retain or modify them based on whether they serve a clear purpose within the paper rather than judging difficulty in isolation.": 10
  },
  "A participant solves a difficult mechanics problem using a method that was not anticipated by the problem setters. Their solution is significantly shorter than the official solution but appears mathematically valid. What should the grading team do?": {
    "Award marks according to the official marking scheme unless the alternative method explicitly demonstrates the same intermediate reasoning.": 40,
    "Verify the alternative method independently and, if valid, accept it and update the marking guidance so mathematically equivalent approaches receive appropriate credit.": 30,
    "Award substantial partial credit for reaching the correct result, while referring the solution to the problem setters before awarding full marks.": 20,
    "Ask the participant to explain their method orally before deciding whether it demonstrates sufficient understanding to receive full credit.": 10
  },
  "Ten minutes before an experimental round, you discover that several of the required sensors are producing inconsistent readings. The round is about to begin and there is no time to replace all of them. What is the best response?": {
    "Start the round as scheduled, instructing teams to repeat measurements where readings appear inconsistent and allowing graders to account for the resulting variation.": 40,
    "Delay the round briefly to test all available sensors, identify the faulty units, and distribute the functioning equipment as evenly as possible before beginning.": 30,
    "Replace the affected measurement with a theoretical value provided to all teams, allowing the experimental component to proceed using the remaining apparatus.": 20,
    "Rapidly determine the extent of the equipment failure, isolate unreliable units, deploy any tested alternatives, and establish one standardized procedure for all affected teams so that the experimental conditions remain as fair as possible.": 10
  },
  "A 45-minute experimental round is running 12 minutes behind schedule because equipment distribution took longer than expected. You still have two rounds planned afterward. What should the event team do?": {
    "Shorten each remaining round slightly so that the overall event still ends close to the original scheduled time, while preserving the order of activities.": 40,
    "Allow the experimental round to continue for its planned duration, then recover the lost time by compressing breaks and transitions between the remaining rounds.": 30,
    "Cancel or significantly shorten the least important later activity so that the core competitive rounds retain their intended duration.": 20,
    "Identify which parts of the remaining schedule are genuinely time-critical, preserve the integrity of the main competitive activities, adjust transitions or lower-priority components where possible, and communicate the revised schedule clearly to all relevant teams.": 10
  },
  "During a major round, you simultaneously receive reports that one room has insufficient answer sheets, another room has a malfunctioning projector, and a participant claims a question is incorrect. You are one of the available department members. What do you do?": {
    "Address the disputed question first because an academic error could affect every participant, while asking another team member to handle the answer-sheet shortage.": 40,
    "Contact the Event Head with all three issues and wait for instructions so that no action is taken without authorization.": 30,
    "Personally resolve whichever issue appears most urgent, then move sequentially through the remaining problems so that each receives your direct attention.": 20,
    "Quickly assess which problems affect fairness or the continuation of the round, delegate solvable logistical issues to available members, and establish a clear process for verifying the disputed question before communicating any correction.": 10
  },
  "You are designing a 20-minute experimental physics challenge for teams of 3–4 high-school students. The activity should test experimental design, data collection, uncertainty analysis, and physical reasoning. You have a limited supply of inexpensive, reusable equipment. Which approach is strongest?": {
    "Give each team a standardized apparatus and procedure, requiring them to collect several measurements, calculate the relevant quantity, and report their experimental uncertainty.": 40,
    "Give teams the same basic apparatus but allow them to choose their measurement procedure, with marks primarily based on the accuracy of their final value and the quality of their uncertainty calculation.": 30,
    "Give different teams slightly different apparatus and ask them to independently determine an unknown quantity, rewarding successful measurements while allowing teams to choose their own procedure and analysis.": 20,
    "Give teams a constrained set of equipment and a clearly defined physical objective, but leave the measurement strategy open; assess the quality of their experimental design, data, uncertainty analysis, interpretation, and final conclusion rather than simply rewarding the closest numerical answer.": 10
  },
  "Two members of the problem-setting team strongly disagree about a proposed physics problem. One argues that it is too difficult for the intended participants, while the other believes it is an appropriate challenge. Both have reasonable arguments and neither is willing to change their position. What is the best way to resolve the disagreement?": {
    "Ask the Event Head to make the final decision, since continued disagreement risks delaying the paper.": 40,
    "Have both members independently explain what assumptions they are making about participant ability, then test the problem with several people who match the target participant level before deciding.": 30,
    "Keep the problem but reduce its marks so that it can function as a high-difficulty discriminator without having too large an impact on the overall score.": 20,
    "Define the intended difficulty and learning objectives first, have independent test-solvers attempt the problem under realistic time constraints, and use their performance and feedback alongside the setters' arguments to make an evidence-based decision.": 10
  },
  "Your team has been asked to introduce an astrophysics component, but you have limited access to telescopes and cannot rely on clear weather. Which concept is strongest for a competitive event?": {
    "Give teams printed star charts and observational clues and ask them to identify constellations and celestial objects, with additional questions testing their knowledge of the objects identified.": 40,
    "Organize an observation-based round using telescopes when conditions permit, while preparing a theoretical astronomy quiz as a backup in case weather prevents observations.": 30,
    "Provide teams with astronomical images and datasets and ask them to identify objects and extract basic properties, with marks awarded for both identification and calculations.": 20,
    "Provide simulated observational data, images, and relevant physical information and ask teams to infer quantities or properties such as stellar classification, distance, luminosity, or motion, making the activity competitive without depending on equipment or weather.": 10
  },
  "You have 30 teams participating in an experimental round, but there are only five identical apparatus setups available. Each team requires approximately 15 minutes to complete the experiment. Which approach is most effective?": {
    "Divide teams into groups and rotate them through the five stations, allowing teams that finish early to begin analyzing their data while they wait for the next available station.": 40,
    "Build several simplified versions of the apparatus so that more teams can work simultaneously, even if the measurements are not perfectly identical across setups.": 30,
    "Allow teams to choose when they want to use an apparatus during a longer open period, giving them flexibility while minimizing idle equipment.": 20,
    "Create timed rotations with standardized setup and reset procedures, assign teams to stations in advance, and ensure that equivalent experimental conditions and timing are maintained across groups.": 10
  },
  "You have three weeks before the event. You are simultaneously responsible for submitting five problem drafts, testing an experimental setup, preparing material lists, and completing an unrelated university deadline that falls in the same week. You realize that completing everything at the last minute would create a significant risk of errors. What is the best approach?": {
    "Complete the problem drafts first because they have the earliest departmental deadline, then shift attention to the experiment and logistics while reserving time near the end for the university deadline.": 40,
    "Break each responsibility into smaller tasks, set earlier personal deadlines, and maintain progress across the different deliverables rather than completing one responsibility entirely before starting another.": 30,
    "Prioritize the event responsibilities for the next two weeks, since they involve coordination with other team members, while setting aside specific time closer to the deadline to complete the university work.": 20,
    "Plan the three weeks around deadlines and task dependencies, giving earlier attention to work that could delay other members or reveal problems in the experimental setup, while leaving enough flexibility to handle the university deadline.": 10
  },
  "One week before the event, you discover that an experimental apparatus you planned to use is unreliable and cannot consistently produce usable data. The experiment has already been included in the schedule and promotional material. What should you do?": {
    "Continue working on the original apparatus for several days, since replacing the experiment this close to the event could create unnecessary disruption, and decide on a backup once you know whether the repairs succeed.": 40,
    "Switch to a simpler experiment using equipment that is already reliable, even if this means changing some of the activity's original learning objectives and format.": 30,
    "Keep the original experiment as planned but prepare instructions for a theoretical version that can be used if the apparatus fails during the event, avoiding major changes to the advertised schedule.": 20,
    "Set a clear testing deadline for the original setup; if it cannot meet the required reliability by then, move to a simpler tested version that retains the core experimental objective rather than continuing to depend on uncertain repairs.": 10
  },
  "You are working on an experimental activity with three other team members. One member has repeatedly missed their assigned tasks and has not responded to messages, leaving the rest of the team to pick up their work. The event is approaching and some of their responsibilities are becoming time-sensitive. What is the best approach?": {
    "Take over their remaining tasks yourself so that the work is completed on time, and discuss the issue with them after the event.": 40,
    "Message them directly, explain the impact of the missed work, and ask them to complete their pending responsibilities by a specific deadline while you continue with your own tasks.": 30,
    "Reassign their tasks among the other members immediately and inform the team lead that the member has not been contributing, since the priority is ensuring that the event preparation stays on track.": 20,
    "Speak to the member directly to understand the reason for the missed work, agree on specific responsibilities and deadlines they can realistically meet, and involve the team lead if the problem continues or begins affecting critical deliverables.": 10
  },
  "During a live physics round, a participant approaches you after seeing their score and argues that a question was ambiguous. They claim that their interpretation is physically valid, while the official marking scheme assumes a different interpretation. The round has already ended. What should you do?": {
    "Explain that the marking scheme was finalized before the event and therefore the score should remain unchanged unless the Event Head decides otherwise.": 40,
    "Hear the participant's reasoning, compare it with the wording of the question and intended solution, and consult the relevant problem setter before making any change to the score.": 30,
    "Ask the participant to submit their argument for review later so that the event schedule is not disrupted, while continuing to use the existing marking scheme for now.": 20,
    "Review whether the participant's interpretation is physically and logically supported by the question as written, compare it with the intended solution and marking scheme, and apply any necessary correction consistently to all affected participants.": 10
  },
  "During testing of an experimental round, two members obtain noticeably different results using the same apparatus. One argues that the unusual result should simply be discarded as an outlier, while the other believes it may indicate a systematic issue with the setup. What should the team do?": {
    "Repeat the measurement several times and use the average of the results to reduce the effect of the unusual measurement.": 40,
    "Compare both sets of measurements and check whether differences in procedure, apparatus handling, or calculation could explain the discrepancy before deciding how to proceed.": 30,
    "Keep the result that is closer to the theoretical value, since this is more likely to represent what the experiment was intended to produce.": 20,
    "Examine the raw data, measurement procedure, uncertainty, and possible systematic effects before deciding whether the result is an outlier, and use the findings to determine whether the setup or instructions need modification.": 10
  },
  "A team member proposes an ambitious experimental activity that would be engaging and scientifically interesting, but it requires some new equipment, additional materials, and more setup and reset time than your usual activities. The event has a fixed budget and multiple teams will participate. What is the best response?": {
    "Test the activity with the available equipment and estimate its material cost and reset time before deciding whether it is practical for the event.": 40,
    "Reject the proposal and use a simpler activity with existing equipment, since reliability, cost, and ease of execution should take priority close to the event.": 30,
    "Keep the main idea but modify the setup to reduce material costs and simplify the reset process while preserving the activity's main experimental objective.": 20,
    "Compare the activity with the team's budget, available resources, preparation time, and expected number of teams before deciding whether the original design or a modified version is more suitable.": 10
  },
  "A balloon-powered boat moves forward when air escapes from the balloon. Which explanation best describes why this happens?": {
    "Air is pushed backward out of the balloon, producing an equal and opposite reaction that propels the boat forward.": 40,
    "The compressed air escaping from the balloon creates a force that is transferred through the boat and causes it to move forward.": 30,
    "The escaping air reduces the pressure around the boat, allowing the surrounding water to push it forward.": 20,
    "As air escapes, the balloon becomes lighter, causing the boat to become lighter and accelerate forward.": 10
  },
  "You find a team using their phones to aid their engineering process. What's the best course of action:": {
    "Subtly but firmly warm them that they have received one warning after which they will lose marks.": 40,
    "Call a head and ask them to make a decision.": 30,
    "Tell them they are disqualified.": 20,
    "Scold them publicly and give them a warning.": 10
  },
  "A catapult is firing but the projectile isn't going far enough, what's the issue?": {
    "The elastic isn't stretched enough to generate enough force.": 40,
    "The elastic is faulty.": 30,
    "The design is incorrect.": 20,
    "The structure is weak.": 10
  },
  "During your category, you notice a delegate is being repeatedly harassed by another delegate. What is the best course of action?": {
    "Check in with the delegate privately and inform your event heads so they inform the respective person in charge so the situation can be handled through the appropriate procedure.": 40,
    "Calmly intervene, make sure the delegate feels safe and inform event heads, while keeping an eye on the situation to make sure it does not continue.": 30,
    "Ask the person involved to stop and continue monitoring the situation, involving the EO only if the behaviour happens again.": 20,
    "Avoid getting involved unless the delegate directly complains, since stepping in without a complaint could make the situation more complicated.": 10
  },
  "You notice your fellow team member is having a heated discussion with a delegate about a rule. As a team manager, what should you do?": {
    "Listen to both sides calmly, ask your team member to take a breather, and direct the delegate to discuss the matter with the Event Head or review the instruction manual. If the situation escalates, contact the respective EO.": 40,
    "Step in calmly, remind both sides to keep the discussion respectful, and suggest checking the instruction manual or consulting the Event Head if the disagreement continues.": 30,
    "Ask your team member to step away from the discussion and handle the delegate yourself so the situation does not escalate further.": 20,
    "Support your team member and ask the delegate to stop arguing, since allowing delegates to challenge volunteers could undermine the team's authority.": 10
  },
  "Your thermocol boat is placed in water but sinks instead of floating. What is the most likely explanation?": {
    "The boat's weight is too large relative to the amount of water it displaces, so the upward buoyant force is insufficient to support it.": 40,
    "The boat may have been constructed in a way that allows water to enter or may have an uneven weight distribution affecting its buoyancy.": 30,
    "The boat may be sitting too low in the water because its design creates too much resistance against the surrounding water.": 20,
    "The balloon may not be producing enough thrust to keep the boat above the water.": 10
  },
  "A delegate asks you for a hint because another team apparently received one from a different volunteer. What is the best response?": {
    "Check the rules and consult the category leadership about the inconsistency before deciding whether the same assistance should be provided to all teams.": 40,
    "Give the delegate the same hint the other team reportedly received so that both teams have access to the same information.": 30,
    "Explain that you cannot provide a hint yourself and ask them to continue according to the instructions unless the Event Head says otherwise.": 20,
    "Give the delegate a different hint of similar difficulty so that neither team receives exactly the same advantage.": 10
  },
  "You see 2 Delegates getting into a heated argument over materials that seems to be on the verge of becoming a physical confrontation, what's the best course of action?": {
    "Get in the middle with minimal physical contact and try to calm them down whilst asking a fellow member to contact a head or phone calling a head.": 40,
    "Try and call an event head to the scene as soon as possible.": 30,
    "Physically apprehend both to prevent anyone getting hurt and to put an end to the situation quickly without interrupting the event progress.": 20,
    "Sit tight and wait for event heads to intervene.": 10
  },
  "You notice a team has a faulty elastic that doesn't allow their catapult to fire. Changing it for them would allow others to demand material replacements that aren’t deserved. What would you do?": {
    "Have it replaced.": 40,
    "You ask the heads to investigate and have it changed if the need is genuine.": 30,
    "Pity their luck and feel bad for them.": 20,
    "You replace it and if anyone else tries to take advantage of the situation you reprimand them.": 10
  },
  "You notice a fellow team member trying to favor his former school during testing by giving them undue favors, what do you do?": {
    "Alert the heads in secret so that they may do whatever seems to be the best course of action.": 40,
    "Both A and C.": 30,
    "Lecture your fellow team members in private and remind them of the integrity of the event.": 20,
    "Lecture them publicly to create an example.": 10
  },
  "You are testing two boat designs made from the same amount of thermocol. Boat A has a wide, flat base while Boat B is narrow and tall. Which design would generally be more stable on water?": {
    "Boat A, because its wider base generally provides greater stability and makes it less likely to tip.": 40,
    "Boat B, because its greater height allows it to displace more water.": 30,
    "Both will have exactly the same stability because they contain the same amount of thermocol.": 20,
    "Boat B, because a taller structure automatically has a lower centre of gravity.": 10
  },
  "A university replaces the roofs of three buildings with green roofs. Six months later, the roofs are absorbing less heat than expected. Your team has one week to decide what to change. You find that: ● Roof A has very healthy plants but gets almost no wind. ● Roof B has poor plant growth but receives strong wind. ● Roof C looks unhealthy, but its temperature is closest to the target. You can only make one change to one roof. What do you choose?": {
    "Leave the roofs unchanged and repeat the measurements under different weather conditions before modifying anything.": 40,
    "Change the plants on Roof A to varieties that tolerate warmer, less ventilated conditions.": 30,
    "Change the growing medium on Roof B to retain more water.": 20,
    "Change the structure around Roof C so that more air can move across its surface.": 10
  },
  "You're given a simple game with one rule: You must make every decision using only information that was available to you before the game began. After playing for a while, you realize that following the rule is making you consistently lose. You cannot change the rule. What do you do?": {
    "Work backwards from your losses to identify what your current strategy is missing.": 40,
    "Try deliberately making a few decisions that seem unreasonable and see whether they produce a different pattern.": 30,
    "Continue following the rule and try to find a better strategy using the information you already have.": 20,
    "Stop trying to win and try to understand why the rule was designed this way.": 10
  },
  "A town has a system that collects and processes 83% of its household waste. Nobody knows what happens to the remaining 17%. Your team is given six months to deal with the problem. There is enough money to pursue only one of these approaches.": {
    "Study the 17% itself and determine what materials make up the missing waste.": 40,
    "Study households that produce unusually little waste and see whether their practices can be reproduced elsewhere.": 30,
    "Study the collection system and identify where waste could be disappearing between households and processing facilities.": 20,
    "Ignore the 17% initially and try to reduce the total amount of waste entering the system.": 10
  },
  "Suppose you are leading a round and teams are performing an activity, you caught a team cheating and using AI, what will you do and how will you handle the situation at that moment?": {
    "You'll just inform the problem to your society higher ups": 40,
    "Giving them a second chance": 30,
    "Ignoring the problem completely": 20,
    "Straightforward eliminating the team from the event": 10
  },
  "Suppose you are leading a round, a team is working in the chemistry lab and they mistakenly broke an apparatus and a toxic chemical spills, what will you do and explain how you will manage the situation?": {
    "Inform the Lab incharge and higher ups immediately and ask them to wash themselves quickly": 40,
    "Ask them to wash themselves quickly and clean the spill with clean water": 30,
    "Clean the spill with a tissue paper and distilled water": 20,
    "Fine the team for breaking the apparatus": 10
  },
  "A team consists of some delegates that were juniors in your school, they reach out to you and ask for bonus marks or try to bribe you to help them in winning the event, what will you do in this situation?": {
    "Explain them that it is unfair to other teams and against your society values": 40,
    "Ignore their messages completely and tell them you did not have time to check your mobile phone": 30,
    "Argue with them and tell them you will try it if possible": 20,
    "Give them a few bonus marks because it is alright in friendships": 10
  },
  "A factory wants to reduce the amount of pollution entering a nearby river. It installs a new treatment system, and laboratory tests show that the concentration of the targeted pollutant in the wastewater has decreased significantly. However, six months later, the river ecosystem appears to be getting worse rather than better. Your team has limited time and funding to investigate. What do you do first?": {
    "Investigate whether reducing the targeted pollutant may have changed other properties of the wastewater or ecosystem that are contributing to the decline.": 40,
    "Compare the river's condition before and after the treatment system was installed to identify any other environmental changes occurring at the same time.": 30,
    "Improve the treatment system further to remove an even greater percentage of the targeted pollutant.": 20,
    "Assume that the ecosystem needs more time to recover and continue operating the system without further investigation.": 10
  },
  "A community wants to produce clean drinking water, but conventional treatment technologies are too expensive to install and maintain. Your team has three months and enough funding to develop only one initial approach. What do you choose?": {
    "Investigate locally available materials, natural processes, and existing infrastructure to identify an unconventional treatment approach that can be tested and adapted to the community.": 40,
    "Study low-cost water treatment technologies used successfully in similar communities and adapt the most suitable one.": 30,
    "Design a simplified version of a conventional treatment system using cheaper materials.": 20,
    "Recommend waiting until sufficient funding becomes available for a conventional treatment plant.": 10
  },
  "A city notices that electricity consumption has increased significantly over the past five years. The city government assumes that population growth is the primary cause and plans to build additional power infrastructure. Before committing to the plan, your team is asked to investigate the problem. What is the most effective approach?": {
    "Examine whether population growth actually explains the increase by investigating other possible factors, such as changes in consumption patterns, technology, industry, infrastructure, or energy efficiency.": 40,
    "Compare electricity consumption trends in similar cities with similar population growth to see whether the same pattern occurs.": 30,
    "Calculate future electricity demand based primarily on projected population growth and recommend infrastructure accordingly.": 20,
    "Accept the government's explanation and focus on determining how quickly new power infrastructure can be built.": 10
  },
  "Linked Genes: Reading the Whole Prediction Genes A and B are linked. An individual with arrangement AB/ab is test-crossed with ab/ab. The recombination frequency is 18%. Assume equal viability of all offspring. Every statement below is scientifically correct. Which gives the MOST complete and useful prediction of the offspring pattern?": {
    "The two parental phenotypes together should account for about 82% of the offspring, while the two recombinant phenotypes together should account for about 18%.": 40,
    "AB/ab and ab/ab should each occur at about 41%, while Ab/ab and aB/ab should each occur at about 9%.": 32,
    "A parental-type offspring is expected to be much more common than a recombinant-type offspring because most gametes are non-recombinant.": 25,
    "AB/ab and ab/ab should tie as the two most frequent individual offspring classes.": 18,
    "The recombinant classes should be the minority because the recombination frequency is well below 50%.": 10
  },
  "Oxygen Transport: Active Tissue Actively respiring muscle has more CO₂, a lower pH and a higher temperature than blood in the lungs. Every statement below is correct under these conditions. Which statement BEST connects these changes to the physiological purpose of haemoglobin in the tissue?": {
    "At the same partial pressure of oxygen, haemoglobin will tend to have a lower percentage saturation in the active tissue than under lung-like conditions.": 40,
    "The dissociation curve shifts to the right, reflecting a reduction in haemoglobin’s affinity for oxygen.": 32,
    "The Bohr effect makes it easier for oxygen to unload where respiration is producing more carbon dioxide and hydrogen ions.": 25,
    "The combined effect of increased CO₂, reduced pH and higher temperature lowers oxygen affinity and promotes oxygen delivery to the metabolically active tissue.": 18,
    "If haemoglobin concentration is unchanged, these conditions mainly alter oxygen affinity and unloading rather than the maximum amount of oxygen that haemoglobin can bind when fully saturated.": 10
  },
  "Same Genome, Different Cell A neuron and a skeletal-muscle cell from the same person contain essentially the same genome but have very different structures and functions. All of the following statements are correct. Which is the BEST overall explanation of how such differences are produced and maintained?": {
    "Different cell types maintain different patterns of gene activity through combinations of transcription factors, regulatory DNA, chromatin state and other layers of gene regulation.": 40,
    "Alternative RNA processing can allow the same gene to contribute different RNA or protein products in different cell types.": 32,
    "Epigenetic marks and chromatin organisation can help stabilise cell-type-specific patterns of gene expression across cell divisions.": 25,
    "Post-transcriptional and translational regulation can make protein abundance differ even when two cells contain transcripts from some of the same genes.": 18,
    "Cell specialisation depends less on having different genes than on using shared genetic information in different combinations, amounts and contexts.": 10
  },
  "Apoptosis: Strength of Evidence A compound reduces the number of cultured cancer cells after 24 hours. Several follow-up findings are obtained, and every finding below is compatible with apoptosis. Which would provide the STRONGEST evidence that apoptosis, rather than another form of cell loss, is the principal mechanism?": {
    "Microscopy shows cell shrinkage, chromatin condensation and membrane blebbing in many treated cells.": 40,
    "Treated cells show externalisation of phosphatidylserine before widespread loss of plasma-membrane integrity.": 32,
    "A DNA-fragmentation assay becomes strongly positive after treatment.": 25,
    "Mitochondrial cytochrome c release is detected soon after exposure to the compound.": 18,
    "Two independent assays show executioner-caspase activation together with early phosphatidylserine externalisation, while membrane integrity is initially preserved.": 10
  },
  "Association, Confounding and Causation Plants of the same species growing beside a busy road have less chlorophyll than plants sampled 2 km away. Every evaluation below is scientifically valid. Which is the BEST evaluation of the claim that vehicle pollution caused the difference?": {
    "Using the same species reduces one obvious biological source of variation, so the comparison is stronger than comparing unrelated species.": 40,
    "A larger sample would improve precision and reduce the influence of random sampling variation, although it would not by itself remove systematic site differences.": 32,
    "The observation supports an association, but a stronger causal claim requires controlling or measuring plausible alternatives such as light, water, nutrients, leaf age and disease.": 25,
    "Statistical significance could strengthen confidence that the observed difference is not just random sampling noise, but it would not by itself establish causation.": 18,
    "The roadside pattern is biologically useful evidence for generating or supporting a pollution hypothesis even though an observational comparison cannot isolate every causal factor.": 10
  },
  "Marking an Equivalent Scientific Answer A marking key expects “water moves into the cell by osmosis.” A participant instead writes that water moves across the partially permeable membrane from higher water potential to lower water potential. The answer is scientifically equivalent, and several scripts have already been marked. Every response below is defensible. Which is the BEST response?": {
    "Accept the scientifically equivalent wording for the current script and record the issue so the key can be clarified before the next event.": 40,
    "Pause final scoring of that item long enough for a second marker to confirm equivalence, then apply the same interpretation to all affected scripts.": 32,
    "Accept the answer now and instruct remaining markers to accept the same wording, then audit earlier scripts if the result could affect qualification.": 25,
    "Preserve the participant’s response and flag the item for moderation before final results, rather than making an irreversible ruling at the desk.": 18,
    "Treat the wording as correct because the concept, not the exact term, is being tested, and re-check all already-marked responses for scientifically equivalent formulations before finalising scores.": 10
  },
  "Correcting an Organiser-Caused Disadvantage In a timed station you tell one team to begin at Question 4 instead of Question 1. The questions can be attempted in any order, but the team loses about two minutes trying to understand the unexpected sequence before working normally. Every action below could form part of a reasonable remedy. Which is the BEST immediate response?": {
    "Correct the instruction, let the team continue, and note the incident so any effect can be considered during moderation.": 40,
    "Correct the instruction and give back the full seven minutes that elapsed before you noticed, choosing to over-compensate rather than risk under-compensating.": 32,
    "Correct the instruction, explain briefly that question order is flexible, and allow the team to continue from whichever question they prefer.": 25,
    "Correct the instruction immediately, restore approximately the time actually lost because of the confusion, document the adjustment, and avoid compensating for time in which useful work was still completed.": 18,
    "Correct the instruction and preserve the team’s existing work, while arranging a later review of whether any question-specific disadvantage remained.": 10
  },
  "Friendship, Honesty and Operational Recovery Your close friend forgot to print an answer sheet and asks you to say the printer malfunctioned if anyone asks. Ten minutes remain before delegates arrive, and the sheets can still be printed if both of you pause a non-urgent task. Every option below contains a reasonable element. Which is the BEST overall response?": {
    "Refuse the false explanation, focus first on getting the sheets printed, and let your friend own the mistake if an explanation is actually needed.": 40,
    "Ask your friend to start printing immediately while you cover the paused task, then discuss the attempted excuse privately once the operational risk has passed.": 32,
    "Print the sheets together, make a brief factual note of the near-miss for the post-event debrief, and avoid turning a solved problem into a public confrontation.": 25,
    "Tell your friend that honesty is non-negotiable, ask them to notify the relevant team member of the delay, and help them finish the printing before delegates arrive.": 18,
    "Refuse to lie, repair the problem jointly, keep any explanation factual and proportionate, and ensure the missed responsibility is addressed afterward rather than hidden or dramatised during the crisis.": 10
  },
  "Bioinformatics Station: Internet Failure A sequence-analysis station depends on an online tool. Twelve teams complete it normally before the internet fails; the remaining teams have not started. You have a pre-tested offline packet containing the same query sequences, saved search outputs and the same scoring rubric. Every response below is defensible. Which is the BEST way to preserve both fairness and the learning objective?": {
    "Use the offline packet for the remaining teams and document the change, because the biological reasoning and scoring targets remain the same even though the interface differs.": 40,
    "Pause the station briefly, verify that the saved outputs exactly correspond to the intended queries, then switch remaining teams to the offline packet under the same time and scoring rules.": 32,
    "Allow the first twelve teams to keep their results but exclude interface-speed elements from scoring if those elements cannot be reproduced offline.": 25,
    "Convert the station for the remaining teams into interpretation of the saved outputs rather than live searching, while keeping only learning objectives common to both versions in the score.": 18,
    "Use the verified offline packet, neutralise any marks that depended specifically on live-tool navigation, preserve the common biological-analysis marks, and record the contingency so every team is compared only on equivalent evidence.": 10
  },
  "Correcting Misinformation Without Undermining a Teammate During a live station, a junior organiser gives a delegate a technically inaccurate explanation. The error is unlikely to affect the delegate’s score, but leaving it uncorrected would teach the biology incorrectly. All five actions below are appropriate at some stage. Which should take PRIORITY?": {
    "After the interaction, give the organiser a concise private explanation of what was inaccurate and how to explain it correctly next time.": 40,
    "For the rest of the session, stay nearby and quietly support the organiser on technical questions until their confidence and accuracy improve.": 32,
    "Briefly and neutrally correct the scientific point for the delegate in the moment, without framing the organiser as incompetent, then debrief the organiser privately afterward.": 25,
    "Give the organiser a short reference sheet or approved explanation for the concept so the same error is less likely to recur.": 18,
    "If the same error repeats after feedback, temporarily shift the organiser toward a role that does not require independent technical explanation.": 10
  },
  "Timing Fault With an Audit Trail After a 30-minute MCQ eliminator, you discover that one team’s timer remained open for 35 minutes. The platform log shows that the team submitted all answers by minute 28 and made no changes afterward. All responses below are reasonable parts of quality control. Which deserves the HIGHEST priority before results are finalised?": {
    "Document the timer fault and correct the system configuration before the next round so the defect cannot recur.": 40,
    "Independently verify the platform log and, if it confirms no answers were changed or newly submitted after minute 30, retain the team’s score without a competitive penalty.": 32,
    "Check whether the platform allowed any hidden advantage after submission, such as viewing feedback or reopening answers, before treating the audit trail as conclusive.": 25,
    "Keep a record of the irregularity with the score so the decision can be explained transparently if another team later asks about timing consistency.": 18,
    "Review whether any element of the round rewarded speed itself; if it did, ensure the timer fault did not affect that component even though answers were submitted before the deadline.": 10
  },
  "While testing a newly built drone on the bench, one motor gets noticeably hotter than the other three within 30 seconds of spinning without propellers. How do you investigate?": {
    "Swap the warm motor with a cool motor on another arm to see if the heat issue follows the motor or stays with the ESC/wiring.": 40,
    "Increase the motor timing in the configuration software to force the motor to run smoother.": 30,
    "Assume it's a defective motor and immediately submit a request to purchase a replacement.": 20,
    "Ignore it for now and attach propellers to see if airflow during flight cools the motor down naturally.": 10
  },
  "You are writing the post project documentation for a drone build that suffered three major crashes during development. How do you frame these failures in the final report?": {
    "Detail the root cause of each crash, the repair process, and the design modifications that can be made to prevent them in the future.": 40,
    "Briefly mention the crashes but focus more on the work that was put in as compared to previous designs and try to frame it as a learning curve": 30,
    "Mention the crashes as a joke to make the report entertaining.": 20,
    "Omit the crashes entirely to make the project look like a complete success.": 10
  },
  "A brushless motor runs significantly hotter than the other three during a 30 second bench test without propellers. How do you start investigating?": {
    "Swap the hot motor with a cool one on another arm to isolate whether it's a motor issue or an ESC/wiring issue.": 40,
    "Connect the drone to configuration software to analyze digital sensor logs and motor timing parameters.": 30,
    "Inspect the physical motor bell, wire solder joints, and mounting screws for mechanical friction or shorts.": 20,
    "Put propellers on and run a brief low altitude flight test to see if ambient airflow cools it down.": 10
  },
  "You notice a potential voltage mismatch in a teammate's wiring diagram just before they start soldering. How do you approach them?": {
    "Ask them to walk you through their wiring plan so you can double-check the voltage specs together.": 40,
    "Hand them the component datasheets and point out the voltage rating difference directly.": 30,
    "Propose testing the circuit with a bench power supply set to low voltage before wiring it permanently.": 20,
    "Bring it up during the next quick team sync so everyone is aligned on the power distribution plan.": 10
  },
  "A drone exhibits a mild drift to the left during hovering. Before adjusting electronic settings or trims, what physical check do you prioritize?": {
    "Check the physical balance and center of gravity by adjusting the position of the battery on the frame.": 40,
    "Check that all four motor mounts are level and that no frame arm is slightly twisted or loose.": 30,
    "Inspect the propellers for subtle bends, cracks, or uneven wear across the four arms.": 20,
    "Verify that the flight controller board is mounted perfectly level on its anti-vibration standoffs.": 10
  },
  "The team is stuck choosing between an Analog video setup (cheaper, lower latency) and a Digital setup (clearer image, higher cost). How do you help break the tie?": {
    "Compare both options strictly against the department's primary goal": 40,
    "Analyze the remaining budget to see which option leaves more safety margin for spare parts and crashes.": 30,
    "Suggest building one primary digital rig and one low-cost analog secondary rig to get the best of both worlds.": 20,
    "Survey the team to see which system members have the most hands-on experience maintaining and fixing.": 10
  },
  "While assembling a complex frame section, you realize you installed an inner plate backward, requiring 30 minutes of teardown to fix. What do you do?": {
    "Pause, document where the instruction or diagram was ambiguous, and then disassemble and fix it.": 40,
    "Immediately take it apart and rebuild it correctly so the mistake doesn't cascade into later steps.": 30,
    "Finish the rest of the frame build first, leaving the correction for the end of the session.": 20,
    "Check if you can adapt or modify the remaining attachments without taking the sub assembly apart.": 10
  },
  "You are assigned to configure a radio link protocol (e.g., ExpressLRS) that you have never used before. What is your learning strategy?": {
    "Read the official documentation and setup guides thoroughly from start to finish before opening the software.": 40,
    "Watch a step by step video tutorial while following along directly with your hardware plugged in.": 30,
    "Ask a senior member or peer who knows the protocol for a quick 10-minute walk through to get you started.": 20,
    "Dive straight into the software, explore the settings menu, and learn by experimenting hands on.": 10
  },
  "Your problem-setting teammate submits 25 original math questions two weeks before the event. What is the most effective vetting procedure before sending the questions to print?": {
    "Have the problem creator write official solution keys and have the whole team check the work via shared document comments.": 40,
    "Run blind test-solving sessions with at least two team members who have never seen the problems, timing them and hunting for alternate interpretations.": 30,
    "Use online plagiarism checkers and AI to search for duplicate questions online.": 20,
    "Print the test directly to keep the problem set confidential and avoid leaks.": 10
  },
  "Fifteen minutes into a 60-minute round, an invigilator confirms that Question 8 has a typo making it unsolvable. What is the immediate, standard course of action?": {
    "Void Question 8 immediately, announce to all rooms to skip it, and award full marks to everyone at the end.": 40,
    "Spend 15 minutes calculating an alternative question, write it on the whiteboards, and add 15 minutes to the overall clock.": 30,
    "Quietly tell only the participants who raised their hands and noticed the typo so the rest of the room isn't disturbed.": 20,
    "Verify the fix centrally, write the exact correction on a central white/blackboard in every hall simultaneously, and log time impact for grading adjustments.": 10
  },
  "A high school competitor uses an advanced university-level theorem to solve a 10-point geometry/algebra problem in two lines. The official rubric planned for a 10-step Euclidean proof. How should this be scored?": {
    "0 points, because the method exceeds the competition syllabus.": 40,
    "Partial credit, because they bypassed the intended problem-solving thought process.": 30,
    "Full credit, provided the theorem was stated accurately, applicable, and mathematically sound.": 20,
    "Disqualify the response and require the student to re-solve it during an interview ( suspicion of AI use.": 10
  },
  "Twenty minutes into a 90-minute high-stakes individual round, a 16-year-old participant begins hyperventilating, crying silently, and stops writing entirely. How should the event team intervene?": {
    "Announce to the room that the round is designed to be difficult to normalize their feelings.": 40,
    "Discreetly signal an available invigilator/teammate to gently guide the student outside into a quiet holding space, offer water, and determine if they feel able to resume with adjusted time if permitted by policy.": 30,
    "Leave the student alone at their desk so they are not embarrassed in front of their peers.": 20,
    "Ask the delegate if they’d like to switch places with another teammate/drop out of the round entirely, with no marks awarded.": 10
  },
  "An audience member shouts out a decisive hint (\"Use Cauchy-Schwarz\") right after a student buzzes in, but before they speak.": {
    "Call a stage freeze, discard the question, issue a direct warning that further audience hints will force an empty-hall policy, and serve a fresh reserve problem to the stage.": 40,
    "Allow the buzzing team to finish their response, but require them to verbally prove and derive the Cauchy-Schwarz step from scratch to verify organic knowledge.": 30,
    "Accept the answer if correct, but penalize the school affiliated with the shouting spectator by deducting points from their delegation score.": 20,
    "Ignore the shout entirely and award the points if the buzzing contestant gives the correct numerical answer within their remaining time.": 10
  },
  "In a live countdown face-off, a participant buzzes and states: \"The answer is 60°.\" The host's solution sheet explicitly requires an angle in radians (pi/3). The student insists their response is mathematically equivalent based on how the host read the prompt.": {
    "Offer the contestant a 5-second window to state the radian equivalent without losing floor control.": 40,
    "Pause the match clock, pull up the written prompt to check if \"in radians\" was explicitly mandated; if not, award full credit for mathematical equivalence.": 30,
    "Rule it incorrect and open the floor to the opposing team.": 20,
    "Mark the answer incorrect on the spot, but file the incident for post-match scoring adjustment in case a tiebreaker is required in round/winner cutoffs.": 10
  },
  "In an auction/point-gamified math round, the leading team intentionally buzzes early on questions they don't know just to exhaust the question timer, taking small negative penalties to run down the clock and protect their lead. Opponents are visibly frustrated.": {
    "Stop the clock and adjust game rules mid-round by requiring teams to write a coherent first step on a whiteboard before buzzing.": 40,
    "Disqualify the leading team immediately for violating the collaborative spirit of PSIFI": 30,
    "Allow the round to conclude under the published rules as written, but initiate a fast review of the penalty structure before the next bracket to scale up negative points for unattempted buzzes.": 20,
    "Directly curbs poor sportsmanship, but injects subjective referee discretion (\"what counts as frivolous?\") into a game round.": 10
  },
  "Two days before the exam print cutoff, your Event Heads get into an escalating argument. Event Head 1 insists on including a highly abstract topology problem that 95% of delegates will fail, while Event Head 2 insists on replacing it with an accessible combinatorics question. They refuse to speak to each other, blocking the final draft.": {
    "Immediately flag the stalemate to higher authority / Academic Officers, laying out the print-deadline risk and requesting an executive tiebreaker so the team doesn’t miss the printing window.": 40,
    "Discreetly approach Event Head 1 privately to persuade them to move their topology problem to an optional \"tiebreaker / bonus\" slot, pitching it as a way to preserve their question without blowing up the paper's base difficulty.": 30,
    "Stay out of the conflict entirely since neither head reports to you, continue formatting the problems that are agreed upon, and wait for them to work it out among themselves before taking further action.": 20,
    "Compile an objective comparison memo cross-referencing both problems against the published study guide, estimated solve times, and historical scoring data, then present it to both leads in a structured manner to aid an objective decision, escalating to the Academic Officers only if the deadlock remains past the deadline.": 10
  },
  "During late-night paper grading, you notice one of your teammates giving generous partial credit to delegates from their own former school, arguing that \"they clearly had the right intuition even if the arithmetic failed.\"": {
    "Verbally reprimand the volunteer at the grading table, manually deduct the extra points on the specific papers, and watch them closely while they grade the rest.": 40,
    "Allow the marks to stand to preserve morale, but balance it by instructing other teammates to be slightly more lenient with competing schools.": 30,
    "Report the teammate to the Event Heads and ask them to recheck the stack for unfair marking.": 20,
    "Relieve the teammate of grading duties, recheck the stack they were assigned and redistribute the work amongst the other teammates, (Assign them another task that doesn't require grading)": 10
  },
  "Right after official scores are posted, an aggressive delegate storms up to you in the round room, slams their paper on your desk, and loudly accuses you of personal bias, claiming your rubric deliberately robbed their school of 1st or Runners up. Other delegates are gathering to watch.": {
    "Keep your tone low and steady, ask them to step to a less crowded area, hear their grievance fully without interruption, and then walk them through the rubric to reassure their worries.": 40,
    "Firmly state that aggressive accusations violate the delegate code of conduct, and tell them you will not review their paper until they speak to you respectfully.": 30,
    "Apologize for the distress and grant a provisional point increase on the spot to defuse the crowd and keep the schedule moving.": 20,
    "Inform your Event Heads about the issue and redirect the delegate’s concerns and the harassment you were subjected to towards them.": 10
  },
  "It is your second week on Flagship A. Your subteam has several things it wants to improve, but nobody has assigned you a specific task yet. You have some free time before the next meeting.": {
    "Review the project and identify an area where you could contribute.": 40,
    "Ask the team lead what needs attention and take their suggested task.": 30,
    "Learn more about the project until someone gives you something specific.": 20,
    "Wait for the next meeting so you avoid overlapping with someone else.": 10
  },
  "Your team is testing a mechanism that has worked several times before. Today it keeps failing in slightly different ways. Several quick fixes have already been attempted without success.": {
    "Ask another member to take over and try solving the problem differently.": 40,
    "Keep making small changes until one of them produces a stable result.": 30,
    "Put the issue aside temporarily and work on another part instead.": 20,
    "Compare the failed attempts and use their differences to plan the next test.": 10
  },
  "During a brainstorming session, someone suggests an unusual way to improve the robot's eventual sorting task. It sounds interesting, but nobody knows whether the current hardware could support it.": {
    "Suggest keeping the idea for later while developing the current system first.": 40,
    "Ask what problem the idea solves and explore a smaller testable version.": 30,
    "Point out the hardware limitations and return to the original approach.": 20,
    "Encourage pursuing the idea immediately because unusual ideas can stand out.": 10
  },
  "Your work depends on something another member was supposed to finish. They have been unavailable for several days, and the deadline is approaching. You have never worked on their part before.": {
    "Wait for them to return because changing their work could create problems.": 40,
    "Tell the team lead that your progress is blocked and await further instructions.": 30,
    "Start rebuilding their part independently so the deadline remains unaffected.": 20,
    "Learn enough about their part to stay productive while keeping everyone informed.": 10
  },
  "Your team is choosing between two possible solutions. Both seem reasonable, but people have started defending their preferred option rather than discussing the actual problem.": {
    "Suggest comparing the assumptions, risks, and expected results of both approaches.": 40,
    "Suggest trying whichever approach appears easier to implement first.": 30,
    "Let the more experienced member choose which approach the team follows.": 20,
    "Stay neutral and let the discussion continue until everyone reaches agreement.": 10
  },
  "Your basic robot arm is finally working reliably. Someone proposes adding an ambitious vision feature before the next demonstration. It could improve the demo, but time is limited.": {
    "Reject the idea for now because the existing system already works.": 40,
    "Start building the feature immediately while the idea is still fresh.": 30,
    "Ask the team to compare the feature's value, effort, and possible risks.": 20,
    "Suggest adding the feature only if the current system stays reliable.": 10
  },
  "Your newly assembled robot arm reaches its target reasonably well when moving slowly. At higher speeds, however, the arm begins shaking and sometimes stops slightly away from the target. Which would be the most useful thing to investigate first?": {
    "Whether the camera is identifying the target from a different position.": 40,
    "Whether the arm structure and joints remain stable during movement.": 30,
    "Whether the gripper needs additional force to reach the target correctly.": 20,
    "Whether the movement algorithm should generate a more complex path.": 10
  },
  "During testing, three joints move approximately as expected, but one joint consistently rotates farther than commanded. The same error appears whenever that joint receives the same command. What would you investigate first?": {
    "Whether the camera is incorrectly detecting the arm's workspace.": 40,
    "Whether the gripper is changing the balance of the entire arm.": 30,
    "Whether that joint has an incorrect calibration or position mapping.": 20,
    "Whether the entire arm requires a different movement sequence.": 10
  },
  "Your arm works normally when one servo moves. When several servos move simultaneously, some become weak or reset unexpectedly. The software appears to be sending the commands correctly. What would you investigate?": {
    "Whether the camera is processing images quickly enough during movement.": 40,
    "Whether the movement algorithm is generating too many commands.": 30,
    "Whether the gripper needs recalibration before testing multiple joints.": 20,
    "Whether the power supply and wiring support simultaneous motor demand.": 10
  },
  "The camera detects a red object reliably in one part of the workspace but struggles when the exact same object is moved elsewhere. The object itself has not changed. Which investigation would be most useful?": {
    "Check whether lighting or background conditions change across the workspace.": 40,
    "Check whether the arm's joints are moving at the correct speed.": 30,
    "Check whether the gripper can hold the object securely in each position.": 20,
    "Check whether another object produces more consistent detection results.": 10
  },
  "You ask the robot to move its gripper to the same location twice. From one starting position it reaches the target smoothly; from another starting position it takes a very different route. What would best explain this difference?": {
    "The camera must be identifying the target differently each time.": 40,
    "The robot can use different joint configurations to reach the same position.": 30,
    "The power supply must change whenever the starting position changes.": 20,
    "The gripper must have a different weight during the second movement.": 10
  },
  "Your team records four trials while testing the arm: Trial Speed Payload Result 1 Low Light Accurate 2 High Light Inaccurate 3 Low Heavy Accurate 4 High Heavy Inaccurate You want to understand what is most likely affecting accuracy.": {
    "Test different payloads while keeping the speed fixed.": 40,
    "Test different speeds while keeping the payload fixed.": 30,
    "Change both speed and payload together to see larger differences.": 20,
    "Repeat the highest-speed trial because it produced the largest error.": 10
  },
  "You joined Flagship A because building a robot sounded exciting. A few weeks later, you spend an afternoon repeating tests, recording results, and recalibrating the same mechanism. What would you most likely do?": {
    "Finish the testing but request more interesting work afterward.": 40,
    "Suggest that experienced members handle repetitive testing whenever possible.": 30,
    "Focus on another task because repetitive work feels less valuable to you.": 20,
    "Complete the testing carefully and look for patterns that improve future testing.": 10
  },
  "How would you describe your approach to teamwork?": {
    "I actively contribute, communicate, and help the team achieve its goals": 40,
    "I complete my assigned tasks and support the team when needed": 30,
    "I prefer working independently but can collaborate when required": 20,
    "I usually wait for others to assign me tasks": 10
  },
  "If you are given multiple tasks with the same deadline, what would you do?": {
    "Prioritize them based on urgency and importance and create a plan": 40,
    "Complete them one by one based on their difficulty": 30,
    "Work on all of them simultaneously": 20,
    "Wait until the deadline is close and then complete them": 10
  },
  "How would you handle a disagreement with a team member?": {
    "Discuss the issue calmly and work toward a solution that benefits the team": 40,
    "Listen to their perspective and then explain my own": 30,
    "Ask team lead to resolve the disagreement": 20,
    "Avoid the disagreement and continue with my own approach": 10
  },
  "What do you think we could do to improve SPADES’ outreach and reach more students?": {
    "Explore new outreach channels, approach different student groups, and suggest creative campaigns": 40,
    "Increase promotion through existing social media and campus channels": 30,
    "Ask other team members to help with outreach": 20,
    "Continue using the existing methods and wait for registrations": 10
  },
  "If you have an important academic deadline and a task due on the same day, what would you do?": {
    "Plan ahead, communicate early, and manage both responsibilities accordingly": 40,
    "Complete the academic work first and then finish the task": 30,
    "Ask a team member for help if necessary": 20,
    "Wait until the deadline approaches before deciding": 10
  },
  "A participant is confused about the registration process and keeps asking the same questions. How would you respond?": {
    "Patiently explain the process and guide them through each step until their registration is complete": 40,
    "Explain the process and share the relevant registration guidelines": 30,
    "Direct them to the registration FAQ or information provided": 20,
    "Ask them to contact someone else for assistance": 10
  },
  "You notice that the information in the registration sheet does not match the details submitted through the registration form. How would you handle the situation?": {
    "Cross-check both sources, contact the participant if needed, and update the record": 40,
    "Confirm the discrepancy with a team member before making any changes": 30,
    "Inform the directors and wait for their instructions before proceeding": 20,
    "Leave the information unchanged until it is addressed later": 10
  },
  "It’s 4:30 PM. A ceremony starts at 6:00 PM. The vendor was supposed to arrive at 3:30 PM but says, “Bas 20 minutes mein pohanch raha hoon.” What do you do?": {
    "Inform your director that the vendor is late and wait for instructions.": 40,
    "Call the vendor, get their exact location and ETA, communicate the delay to your director, identify what can be prepared without them, and keep following up until they arrive.": 30,
    "Assume they’ll arrive soon because there is still 1.5 hours left.": 20,
    "Call the vendor repeatedly and tell them they absolutely need to hurry.": 10
  },
  "You’re given a spreadsheet containing 70 guests, their contact information, PR status, arrival time, and assigned event. You notice several entries are inconsistent or missing. What do you do?": {
    "Only fix the entries you personally need for your assigned task.": 40,
    "Leave it because whoever made the spreadsheet probably knows what they’re doing.": 30,
    "Clean and standardize the sheet, flag missing information, verify what you can, and create a clear way to track unresolved entries.": 20,
    "Highlight the incomplete cells and ask the relevant people to fill them.": 10
  },
  "An artist’s manager quotes you PKR 650,000. Your approved budget is PKR 500,000. What is your first approach?": {
    "Tell them your maximum budget is PKR 500,000 and ask if they can match it.": 40,
    "Tell them another artist has offered PKR 450,000 even though that isn’t true.": 30,
    "Understand what the quote includes, ask about flexibility, negotiate professionally using your event and platform as leverage, and explore adjustments to the package.": 20,
    "Immediately start looking for another artist within budget.": 10
  },
  "You have been standing at a venue supervising setup for three hours. The carpeting is done, but some chairs are misaligned, the backdrop isn’t centered, and the vendor says, “Sir bas theek hai, nobody will notice.” What do you do?": {
    "Let it go because the vendor has more experience with setups.": 40,
    "Fix the most noticeable problems and let the minor ones go.": 30,
    "Politely but firmly have the incorrect elements fixed before signing off on the setup.": 20,
    "Send pictures to your director and ask whether the setup is acceptable.": 10
  },
  "Your team needs three possible concepts for an opening ceremony in two days. You’ve never planned an opening ceremony before. What do you do?": {
    "Search Pinterest, Instagram, and TikTok and send interesting setups you find.": 40,
    "Research other events and previous SPADES events, generate concepts, check their practical feasibility and cost, and present your strongest options.": 30,
    "Ask your director what exact concept they want before starting.": 20,
    "Brainstorm several ideas with friends or team members and present the best ones.": 10
  },
  "During the event, you are carrying refreshments from one building to another when a vendor calls saying they need someone at the concert venue immediately. What do you do?": {
    "Immediately go to the vendor because the concert is more important.": 40,
    "Call your director and wait for them to decide what you should do.": 30,
    "Assess which task is more time-sensitive, coordinate with a teammate to cover one of them, and make sure neither task gets abandoned.": 20,
    "Finish the refreshments task as quickly as possible and then head to the venue.": 10
  },
  "A vendor becomes irritated during a negotiation and says, “Aap log students ho, aapko rates ka idea nahi hai.” What do you do?": {
    "Respond firmly that you know the market and that they’re overcharging you.": 40,
    "Stay calm, ask them to break down the quotation, compare it with your requirements and market information, and continue negotiating professionally.": 30,
    "Tell them you have other quotations and ask them for their best final rate.": 20,
    "End the conversation and find another vendor who is easier to work with.": 10
  },
  "Your director asks, “What’s the status of tomorrow’s closing ceremony setup?” Which response would you ideally be able to give?": {
    "“Everything is mostly confirmed; I’m just checking one or two things.”": 40,
    "“I’ll call everyone right now and update you.”": 30,
    "“Stage confirmed for 4 PM, chairs and tables at 4:30, carpeting confirmed, vendor contact saved, quotation approved, and one lighting item is still pending which I’m following up on.”": 20,
    "“The vendor said everything will be ready.”": 10
  },
  "It’s 11:30 PM during the flagship event. You’ve already been working most of the day. Your assigned event is over, but another Social & Operations setup is understaffed. What do you do?": {
    "Ask whether they genuinely need you because your assigned work is already finished.": 40,
    "Check with the team and help wherever you’re most needed before leaving.": 30,
    "Leave because everyone should be responsible for their own assigned event.": 20,
    "Help for a while but leave once the immediate rush has been handled.": 10
  },
  "You have quotations from three vendors. Vendor A is the cheapest but has been unreliable. Vendor B costs 12% more but has consistently delivered. Vendor C is the most expensive but offers the best-looking setup. Which approach do you take?": {
    "Choose Vendor C because Social events need to look as impressive as possible.": 40,
    "Choose Vendor B because reliability generally justifies a moderate premium.": 30,
    "Compare cost, reliability, quality, event requirements, and available budget before making and justifying a recommendation.": 20,
    "Choose Vendor A because staying under budget should be the priority.": 10
  },
  "You have 45 minutes before an event and suddenly receive five different tasks from different people. What do you do?": {
    "Ask your director which task they want you to do first.": 40,
    "Start with whichever task has the closest deadline and work through them sequentially.": 30,
    "Start with the easiest tasks so you can clear your task list quickly.": 20,
    "Quickly prioritize them by urgency and dependency, identify what only you can do, delegate what can be delegated, and communicate accordingly.": 10
  },
  "You’re told at 5 PM that 150 extra chairs are needed by 7 PM. Your usual vendor says they can’t provide them. What do you do?": {
    "Ask the usual vendor if they know someone who can arrange them while also calling alternatives yourself.": 40,
    "Explain that arranging 150 chairs with two hours’ notice isn’t realistically possible.": 30,
    "Immediately contact alternative vendors and resources, check what can be sourced internally, coordinate transport and logistics, and keep the relevant lead updated.": 20,
    "Tell your director immediately and ask what they want you to do.": 10
  },
  "A senior team member gives you an instruction that you think will cause a logistical problem later. What do you do?": {
    "Follow the instruction because they probably know something you don’t.": 40,
    "Raise the concern respectfully, explain the likely consequence, suggest an alternative, and then follow the final decision.": 30,
    "Do it your own way because you’ll be responsible if it goes wrong.": 20,
    "Ask them whether they’re sure before proceeding.": 10
  },
  "During a high-pressure setup, your director hasn’t replied for 10 minutes and a vendor needs an immediate minor decision. What do you do?": {
    "Keep calling your director until they answer.": 40,
    "Tell the vendor they must wait because you haven’t received authorization.": 30,
    "Ask another experienced department member for their judgment.": 20,
    "If the decision is within your responsibility and is low-risk or reversible, make the best-informed decision and update your director.": 10
  },
  "You discover that you personally made an error in a vendor sheet that could affect tomorrow’s setup. Nobody else has noticed yet. What do you do?": {
    "Correct the sheet immediately and only mention the mistake if it has already affected something.": 40,
    "Ask a teammate what they think you should do before telling anyone.": 30,
    "Correct what you can immediately, inform whoever needs to know, clearly explain the potential impact, and help implement the solution.": 20,
    "Quietly fix it because admitting the mistake unnecessarily could create panic.": 10
  },
  "No budget cap, no deadline, no supervisor. You build one thing for a month. What is it?": {
    "Something that does one small thing beautifully and serves no practical purpose.": 40,
    "A fix for something that annoys you every single day.": 30,
    "A tool the department keeps needing and does not have.": 20,
    "Something built to win the next competition you enter.": 10
  },
  "Rank these by how much they would ruin your week.": {
    "Someone moved your tools and did not say where.": 40,
    "Uncommented code you have to modify tonight.": 30,
    "One unlabeled cable in a bundle of twelve.": 20,
    "A single missing M3 screw.": 10
  },
  "You get 10 thousand rupees and three weeks to build a demo. How do you scope it?": {
    "One feature that works reliably, with enough budget held back to rebuild it once.": 40,
    "Something ambitious with features trimmed as the deadline approaches.": 30,
    "A kit you modify into your own version so you start from a working baseline.": 20,
    "The strongest components you can afford with the scope shaped around what you end up with.": 10
  },
  "You walk into the lab and catch a smell. Rank these by how fast you move.": {
    "Burning insulation.": 40,
    "Hot plastic somewhere nothing should be heating.": 30,
    "Ozone near a power supply.": 20,
    "Solder flux from a bench nobody is sitting at.": 10
  },
  "Your project has to survive being carried across campus in a rickshaw.": {
    "Nothing loose, everything screwed down, every cable given strain relief before it moves, moving parts ziptied.": 40,
    "A padded box and a written plan for reassembling it on arrival.": 30,
    "Broken into subassemblies, each carried separately and joined at the venue.": 20,
    "Carried flat and level, with the driver asked to take it slow.": 10
  },
  "Your team runs four tests on a part that keeps failing. Trial 1: cool, slow. Passed. Trial 2: cool, fast. Passed. Trial 3: warm, slow. Passed. Trial 4: warm, fast. Failed. What do you test next?": {
    "Repeat trial 4 to confirm it fails consistently, then hold it warm and sweep the speed.": 40,
    "Hold the speed fast and sweep temperature, since trial 2 passed at that same speed.": 30,
    "Add trials between the four settings to find where the failure starts.": 20,
    "Move both variables together in smaller steps to reach the boundary faster.": 10
  },
  "It is three in the morning. One bug left. Presentation at nine.": {
    "Message the teammate who is still awake and pair on it.": 40,
    "Sleep three hours and restart at six.": 30,
    "Keep going alone while the whole system is still loaded in your head.": 20,
    "Freeze the bug, build a demo path that routes around it, and present what works.": 10
  },
  "You have been stuck on one problem for a while. When do you ask for help?": {
    "At around forty-five minutes, arriving with a list of everything you have ruled out.": 40,
    "As soon as you are stuck.": 30,
    "At hour three, once you have exhausted your own ideas.": 20,
    "Next time you see the right person, rather than interrupting them for it.": 10
  },
  "A teammate has missed three deadlines.": {
    "Ask them directly what is blocking them, then redistribute based on the answer.": 40,
    "Raise it with the department head before it reaches the final build.": 30,
    "Cover their share through this deadline and talk to them once it is done.": 20,
    "Put it to the whole team so the workload gets rebalanced in the open.": 10
  },
  "Your idea gets rejected in a planning meeting.": {
    "Ask what the specific objection is, then come back with a prototype that answers it.": 40,
    "Back the chosen plan and commit to it fully.": 30,
    "Make the case again in the room, with the reasoning you have.": 20,
    "Build it in your own time and show them once it works.": 10
  },
  "Your subsystem works. Nobody else on the team knows how to use it.": {
    "Document the wiring and comment the code so anyone can pick it up next week.": 40,
    "Walk two teammates through it at the workbench.": 30,
    "Record a video explaining it and share it with the team.": 20,
    "Stay reachable as the person who knows it and answer questions as they come.": 10
  },
  "A build fails two hours before the showcase.": {
    "Strip it back to the smallest version that runs and demo that.": 40,
    "Isolate subsystems one at a time until you find the break.": 30,
    "Hand it to whoever built that section and support them.": 20,
    "Ask the organizers for a later slot and use the extra time.": 10
  },
  "Someone hands you a component you have never touched. What is your first move?": {
    "Read the datasheet, wire a minimal test circuit, and confirm it works in isolation.": 40,
    "Find a tutorial, replicate the wiring exactly, then adapt it once something responds.": 30,
    "Ask a senior member to walk you through it.": 20,
    "Connect it to the build on a current-limited supply and watch what it draws.": 10
  },
  "Your servo jitters every time the DC motor spins up. What do you check first?": {
    "Whether the motor and the logic board share one power rail, browning out the servo on every spike.": 40,
    "Whether decoupling capacitors are in place across the supply.": 30,
    "Whether the servo has failed, swapping in a known-good one to isolate the fault.": 20,
    "Whether your control timing overlaps the motor's PWM.": 10
  },
  "Your robotic arm has to repeat the same motion accurately, fifty times in a row. What matters most?": {
    "A rigid frame with minimal backlash in the joints.": 40,
    "Encoders and closed-loop feedback, so the arm knows where it actually is.": 30,
    "Motors sized with enough headroom that they never stall under load. Acceleration profiles gentle enough to keep the arm from overshooting.": 10
  },
  "How would you rate your current editing ability? 🟡": {
    "Beginner: still learning the basics 🟢": 40,
    "Intermediate: can independently create good content 🔵": 30,
    "Advanced: comfortable with complex edits, effects, sound, and/or color": 20,
    "I know where the transition button is, so basically advanced 🎬": 10
  },
  "Which editing software are you most comfortable using? 🎨": {
    "Premiere Pro / After Effects 📱": 40,
    "DaVinci Resolve 🗿": 30,
    "CapCut / VN / Other": 20,
    "Microwave ✅": 10
  },
  "You receive hundreds of unorganized clips from an event. What’s your first step? 🙏🏽": {
    "Organize and label the footage ✌🏽": 40,
    "Watch through everything and shortlist clips 💀": 30,
    "Start editing on the go": 20,
    "Ctrl+A → Del → problem solved": 10
  },
  "You’re asked to learn a new skill for a specific project that you’re unfamiliar with. What’s your 📚 approach? 🫂": {
    "Learn what I need and give it a shot 🙅🏽": 40,
    "Ask someone experienced to guide me 🫡": 30,
    "Try to avoid it and stick to what I already know": 20,
    "YouTube tutorial at 2× speed. We move": 10
  },
  "You’ve spent hours making a reel, but your directors ask you to change most of it. What do 👍🏽 you do? 🤔": {
    "Take the feedback positively and improve the work 😤": 40,
    "Ask why the changes are needed and then revise accordingly 😭": 30,
    "Defend my original work and only make the changes I agree with": 20,
    "“But I spent FOUR HOURS on this.”": 10
  },
  "SPADES has a workshop tomorrow and needs MnP coverage early morning, but you also 🤙🏽 have an assignment due the same time. What do you do? 🔥": {
    "Plan ahead and make time for both 😤": 40,
    "Let the team know my availability and contribute where I can 🥀": 30,
    "Prioritize the assignment and skip the coverage": 20,
    "Show up to the workshop and hope my instructor understands": 10
  },
  "During a shoot, the approved, pre-planned content idea isn’t working. What do you do? 🫂": {
    "Improvise and suggest a new approach ‼️": 40,
    "Ask the team for ideas and adapt 🙏🏽 😭": 30,
    "Stick to the original plan": 20,
    "Leave it and hope the directors don’t notice": 10
  },
  "You are assigned a deliverable with a team of 3 people, but they all have different ideas for 🥰 the shoot. What do you do? 🤩": {
    "Discuss with them and choose what works best ❓": 40,
    "Push for the idea I think is strongest 😭": 30,
    "Go with whatever the majority prefers": 20,
    "Make all three and let future generations decide 😝": 10
  },
  "Which sounds most like you during a society event? 😛": {
    "I’m usually the one talking to everyone and getting people involved 🤨": 40,
    "I’m comfortable meeting new people but prefer staying focused on the task 😱": 30,
    "I’m more comfortable working behind the scenes": 20,
    "If there’s a camera, I will eventually end up in front of it 🫣": 10
  },
  "What motivates you most to join SPADES M&P? 🫡": {
    "Creating content and improving my creative skills 💥": 40,
    "Meeting people and being part of a fun, active team 👊🏽": 30,
    "Getting hands-on experience with media and event coverage": 20,
    "All of the above — I want to make cool shit with cool people": 10
  },
  "You have to design a mystery where participants receive 6 clues. Which setup would you find most interesting to develop?": {
    "All 6 clues point toward the same suspect, but some are stronger than others.": 40,
    "4 clues point toward one explanation, while 2 subtly support a competing explanation.": 30,
    "Each clue points toward a different suspect, leaving participants to decide which clues matter.": 20,
    "The clues seem unrelated initially but become connected as participants uncover more information.": 10
  },
  "You test your case and realize almost every group solves it within five minutes. What would you change first?": {
    "Add another layer to the reasoning required to connect the existing clues.": 40,
    "Introduce a misleading clue that initially supports an alternative explanation.": 30,
    "Remove some of the more obvious clues and make participants work with less information.": 20,
    "Add an additional stage to the case so solving it requires reaching multiple conclusions.": 10
  },
  "A participant solves your case using a method your team never intended, but their reasoning is internally consistent. What would you be most interested in?": {
    "Whether their answer matches the intended solution.": 40,
    "Whether their reasoning can actually be supported by the evidence provided.": 30,
    "Whether other teams could reasonably arrive at the same conclusion.": 20,
    "Whether their approach exposes an unintended weakness in the case.": 10
  },
  "Halfway through designing a round, you realize your original idea is much harder to execute than expected. Which direction would you lean toward?": {
    "Simplify the original idea while preserving its central concept.": 40,
    "Replace the difficult component with a different mechanism that tests the same skill.": 30,
    "Restructure the round around the parts that are already working.": 20,
    "Keep the difficult component and try to find a way to make it work.": 10
  },
  "You have one final clue to add to a case. Which kind would you choose?": {
    "One that strongly confirms the most likely explanation.": 40,
    "One that makes participants reconsider an assumption they have already made.": 30,
    "One that connects two pieces of evidence that initially seemed unrelated.": 20,
    "One that introduces a new possibility without immediately resolving it.": 10
  },
  "You’re brainstorming a new SCB round from scratch. What would you naturally start with?": {
    "A cool storyline or setting.": 40,
    "A skill or type of thinking you want participants to use.": 30,
    "A unique game mechanic that could make the round memorable.": 20,
    "An interesting scientific or forensic concept that could be incorporated.": 10
  },
  "While designing the case for SCB, what would you priotitize?": {
    "Creating a complex network of information without a clear answer, leaving room for question in all suspects": 40,
    "A simple case that eventually hints strongly towards particular suspect(s)": 30,
    "A complex yet guided case that relies on theatrics and more direct reveals to lead to a more simple answer": 20,
    "A case that runs deep and leaves room for many different avenues of exploration and multiple correct conclusions, following a broadly linear narrative": 10
  },
  "After reviewing what teams are thinking post round, you find that most teams are coming to virtually the same answer, even with differing analysis. How would you respond?": {
    "Keep everything as is and rely on properly reviewing how they reached their conclusion to judge them": 40,
    "Add a few more clues that hint towards a second possibility to diversify answers in an effort to weed out unsure teams": 30,
    "Add more complexity in the final round and judge top teams on the basis of overall deduction skills across all three rounds to add another layer of marking": 20,
    "Increase amount/size of deliverables to evaluate just how much and how deep every team has thought of the case, alongside how many angles they considered.": 10
  },
  "Two people you are collaborating with on the design team have very different visual styles for the same event campaign, and you need one unified identity. What do you do?": {
    "Bring all the team together, identify strong elements from each style, and merge them into one direction with clear reasoning.": 40,
    "Pick the style you personally prefer and explain your decision to the team afterward.": 30,
    "Let all designers keep working separately and try to combine the outputs later.": 20,
    "Escalate the decision entirely to a director and wait for them to decide.": 10
  },
  "The department director shares a design they made themselves for an upcoming event, and you disagree with their design choices, and it doesn't fit the event's tone. They seem happy with it and are about to send it off. What do you do as a team member?": {
    "Say something before it goes out, focused on the work, not the person.": 40,
    "Ask a clarifying question first ) to understand their intent before deciding whether your concern still stands.": 30,
    "Say nothing directly, but mention your concerns to another teammate to see if they feel the same way before anyone raises it.": 20,
    "Stay quiet, since it's the director's call and their name/authority is on the decision either way.": 10
  },
  "You have two tasks due the same day: one is a small, low visibility internal graphic; the other is a flashier social media post that will get a lot of views but isn't due until later that evening. You only have time to give one your full effort right now. Which do you prioritize, and why?": {
    "The internal graphic, visibility isn't the same as importance, and reliability on the boring work builds more trust over time.": 40,
    "The social post, it's later in the day, so you do the internal one adequately now and save your best effort for the piece more people will see.": 30,
    "Split your effort evenly across both rather than fully prioritizing either, aiming to avoid a truly weak version of anything.": 20,
    "Whichever one you personally find more interesting to work on, since motivation affects output quality.": 10
  },
  "You're told a design \"needs to pop more,\" but you personally think the current version is stronger and more tasteful. What's your next move?": {
    "Make the requested change as asked, then separately share a version closer to your original so both can be compared side by side.": 40,
    "Try to understand what \"pop more\" is actually solving for before deciding how to respond.": 30,
    "Adjust the design to add more visual intensity, trusting that the person asking has context you might not.": 20,
    "Keep the design mostly as is, making only minor tweaks, since you believe restraint is the stronger choice here.": 10
  },
  "You receive harsh criticism about a design you're genuinely proud of. How do you respond?": {
    "Separate your emotional reaction from the useful parts of the feedback, ask clarifying questions, and use it to improve the next version.": 40,
    "Feel upset initially, but eventually come around and incorporate the feedback after some reflection.": 30,
    "Defend the design strongly without giving the criticism much real consideration.": 20,
    "Avoid attempting similar designs in the future to sidestep negative feedback.": 10
  },
  "Your team is behind schedule on a major visual campaign and people are stressed. What role do you take?": {
    "Reassess priorities, delegate tasks based on people's strengths, and keep morale up while pushing toward the deadline.": 40,
    "Focus on completing your own tasks well and trust that others will catch up.": 30,
    "Take on other people's work yourself to make sure everything gets done.": 20,
    "Point out who's falling behind to create a sense of urgency.": 10
  },
  "You have an idea for a new design initiative that's outside your usual responsibilities. What do you do?": {
    "Pitch it to your team or department head with a clear plan of what it would take to execute and what resources are needed.": 40,
    "Mention the idea casually to gauge interest before putting together a formal pitch.": 30,
    "Keep it to yourself, since it falls outside your usual role.": 20,
    "Start executing it on your own without discussing it with anyone first.": 10
  },
  "You have to design a poster for an event with the most boring possible topic. How do you make it not put people to sleep?": {
    "Find one genuinely funny or unexpected angle on the topic and build the whole design around that idea.": 40,
    "Keep it clean and minimal and let good typography and whitespace do the work instead of trying to force excitement.": 30,
    "Lean on bold colors and bigger fonts than usual to make it visually loud even if the content stays dry.": 20,
    "Add a meme or trending reference to make it feel more relatable and current.": 10
  },
  "You are approaching a brand for sponsorship for your university’s flagship event. The brand asks, “What’s in it for us?” What should you prioritize in your response?": {
    "Clearly present the brand visibility, target audience, deliverables, and measurable value they will receive in exchange for their sponsorship.": 40,
    "Tell them that many students will attend the event and their brand will get good exposure.": 30,
    "Emphasize that other well-known brands have already sponsored the event.": 20,
    "Offer them a discounted sponsorship package immediately to convince them. ⸻": 10
  },
  "A company cannot provide an in-cash sponsorship but offers products worth approximately PKR 100,000 for your event. What is the most appropriate approach?": {
    "Evaluate the products against your event’s actual needs and negotiate a barter arrangement with clearly defined deliverables and equivalent value.": 40,
    "Accept the products because PKR 100,000 in products is still better than receiving nothing.": 30,
    "Ask them to increase the product quantity before agreeing.": 20,
    "Reject the offer because your department only deals with monetary sponsors. ⸻": 10
  },
  "A potential sponsor says, “Your sponsorship packages are too expensive.” What would be the strongest next step?": {
    "Ask about their budget and objectives, understand their constraints, and explore a customized package that preserves value for both sides.": 40,
    "Immediately reduce the price by 30% to close the deal.": 30,
    "Tell them another brand has agreed to pay the full amount.": 20,
    "Thank them and move on to another sponsor. ⸻": 10
  },
  "You have contacted 30 brands for an event. Ten have opened your proposal, but only two have responded. What should you do?": {
    "Analyze the outreach and follow-up process, personalize follow-ups, refine the value proposition, and prioritize high-potential leads.": 40,
    "Send the exact same message to another 30 brands.": 30,
    "Follow up with all ten brands repeatedly until they respond.": 20,
    "Assume that brands are simply not interested in sponsoring university events. ⸻": 10
  },
  "You are responsible for inviting a high-profile guest speaker for an event. The speaker’s assistant asks for more information before confirming. What should you send first?": {
    "A concise professional brief covering the event, audience, date/time, expected role, topic/theme, format, and relevant organizational details.": 40,
    "A long message explaining everything about your society and all previous events.": 30,
    "Only the event poster and ask whether they are interested.": 20,
    "Ask them to attend first and provide the details after they agree. ⸻": 10
  },
  "A confirmed panelist cancels two days before an event. You have three potential replacements: a highly relevant expert who may be unavailable, a moderately relevant expert who is available, and a famous personality with little connection to the topic. What should you do?": {
    "Contact the highly relevant expert immediately while simultaneously approaching the available relevant candidate as a backup.": 40,
    "Invite the famous personality because their name will attract more attendees.": 30,
    "Wait for the original panelist to reconsider before contacting anyone else.": 20,
    "Cancel the panel because finding a replacement at short notice is difficult. ⸻": 10
  },
  "A sponsor agrees to provide PKR 50,000, but your team later realizes that one of the promised deliverables cannot realistically be provided. What is the best response?": {
    "Communicate the issue proactively, explain why the deliverable cannot be fulfilled, and negotiate an equivalent alternative that maintains the sponsor’s value.": 40,
    "Provide the deliverable anyway, even if it compromises the event.": 30,
    "Remove the deliverable without informing the sponsor.": 20,
    "Wait until after the event and explain the situation if they complain. ⸻": 10
  },
  "You are negotiating with a brand that wants exclusive category rights at your event but is offering significantly less than your expected sponsorship value. What should you do?": {
    "Assess the value of exclusivity, calculate its opportunity cost, and negotiate the sponsorship amount or scope accordingly.": 40,
    "Give them exclusivity because securing one sponsor is better than securing multiple sponsors.": 30,
    "Reject exclusivity immediately without discussing alternatives.": 20,
    "Agree to exclusivity and approach competing brands secretly. ⸻": 10
  },
  "A potential judge for a competition has excellent credentials but has not responded to your first two messages. What is the most professional approach?": {
    "Send a concise, respectful follow-up through the appropriate channel, reiterating the key details and making it easy for them to respond.": 40,
    "Send multiple messages across all their social media accounts.": 30,
    "Ask someone you know to pressure them into responding.": 20,
    "Remove them immediately and never contact them again. ⸻": 10
  },
  "Your Marketing and External Relations teams are both approaching the same company for an event—Marketing for sponsorship and ERA for a guest speaker. What should you do?": {
    "Coordinate internally, establish a single communication strategy, and approach the company with a unified proposal covering both opportunities.": 40,
    "Let both teams continue independently so the company has more opportunities to engage.": 30,
    "Ask the team that contacts the company first to handle everything.": 20,
    "Stop one team from contacting the company without informing them why.": 10
  },
  "You are helping organize a GBM, but attendance has been dropping over the past few sessions. What would you do?": {
    "Look into why attendance is dropping, gather feedback from members, and redesign the GBM around what would actually make people want to attend.": 40,
    "Try introducing more interactive activities and see whether engagement improves in the next few sessions.": 30,
    "Make the GBM shorter and announce it more frequently so people are reminded to attend.": 20,
    "Assume people are simply busy and continue running GBMs the same way.": 10
  },
  "A new member joins SPADES and barely speaks during their first few meetings. They seem interested but aren't participating much. What would you do?": {
    "Check in with them casually, understand how they're settling in, and find a low-pressure way for them to become involved.": 40,
    "Give them a small responsibility during the next activity so they have a natural opportunity to participate.": 30,
    "Encourage them to speak up more during meetings and give them time to adjust.": 20,
    "Assume they're not interested and wait for them to approach HR themselves.": 10
  },
  "A member who used to be extremely active has suddenly become distant and is missing several society tasks. What is your first move?": {
    "Reach out privately, check whether something has changed, and understand the situation before deciding what action is appropriate.": 40,
    "Ask whether they need help or whether their workload needs to be adjusted temporarily.": 30,
    "Inform their department head that the member seems disengaged and let them handle it.": 20,
    "Call them out during the next meeting because everyone has responsibilities to fulfill.": 10
  },
  "You are planning a society hangout, but the HR team has very different ideas about what to do. Everyone is attached to their own idea and planning has stalled. What do you do?": {
    "Identify what the hangout is actually meant to achieve, such as bonding, relaxation, or welcoming new members, and choose an activity that best fits that goal.": 40,
    "Shortlist the strongest ideas and quickly vote or test them with the team.": 30,
    "Combine parts of everyone's ideas into one larger activity so nobody feels ignored.": 20,
    "Let the loudest or most senior person decide so planning can move forward.": 10
  },
  "You are responsible for keeping the HR database updated. You discover that several members have missing or inconsistent information shortly before an important recruitment cycle. What do you do?": {
    "Identify the inconsistencies, verify the information from reliable sources, and establish a clear process to keep the database accurate going forward.": 40,
    "Fix the most important missing information first and create a list of remaining issues to resolve later.": 30,
    "Ask members to update their own information and assume most of the errors will be corrected.": 20,
    "Leave the existing database alone because fixing everything this close to recruitment would take too much time.": 10
  },
  "Another department suddenly needs volunteers for an important event task. Your HR team is already busy with its own work. What do you do?": {
    "Check the team's capacity, identify what can be temporarily deprioritized, and coordinate volunteers without compromising critical HR responsibilities.": 40,
    "Ask available HR members whether they can volunteer and then distribute the extra workload fairly.": 30,
    "Send whoever is free at the moment, even if they aren't particularly suited to the task.": 20,
    "Refuse because HR already has its own responsibilities.": 10
  },
  "You have been asked to make the slides for the next GBM, but the information from different departments is arriving late and in inconsistent formats. The GBM is tomorrow. What do you do?": {
    "Identify the essential information, standardize it into a clear structure, and follow up directly with departments on anything genuinely missing or unclear.": 40,
    "Build the presentation around the information you already have and leave clearly marked spaces for anything still pending.": 30,
    "Ask the HR head to chase every department for their information before you start making the slides.": 20,
    "Put all the information into the slides as received, even if the formatting and details are inconsistent.": 10
  },
  "A member tells you privately that they feel their department doesn't value their contributions. What do you do?": {
    "Listen first and try to understand specific examples, and then work with the relevant people to address the issue constructively.": 40,
    "Ask what kind of recognition or support they feel is missing and see whether something practical can be changed.": 30,
    "Tell their department head that the member seems unhappy and ask them to check in.": 20,
    "Immediately agree with the member and tell them that their department is clearly treating them unfairly.": 10
  },
  "You are asked why you want to join HR. Which answer best reflects what would make someone genuinely suited to the department?": {
    "I like working with people, and I want to help build a society where members feel included, supported, motivated, and able to contribute.": 40,
    "I enjoy organizing activities, communicating with people, and helping things run smoothly behind the scenes.": 30,
    "I think HR would help me become more confident and improve my communication skills.": 20,
    "HR seems like the fun department because you get to hang out with everyone and attend events.": 10
  },
  "You accidentally send an incorrect announcement to the entire society about an upcoming event. You realize the mistake shortly afterward. What do you do?": {
    "Correct the information promptly, and send the corrected information reaches everyone who received the original message.": 40,
    "Inform your HR Head immediately and ask how they would like the correction handled.": 30,
    "Delete the original message if possible and send the correct one without drawing attention to the mistake.": 20,
    "Wait to see whether anyone notices before deciding whether a correction is necessary.": 10
  }
};

function normalizeResponses_(responses) {
  if (!Array.isArray(responses) || responses.length !== 11) {
    throw new Error("Invalid response count");
  }

  return responses.map(function(r) {
    const question = cleanText_(r.question, 1000);
    const selectedAnswer = cleanText_(r.selectedAnswer, 1000);

    if (!question || !selectedAnswer) {
      throw new Error("Invalid response");
    }

    // Reject answers that are not in the server-side question bank.
    if (
      !ANSWER_KEY[question] ||
      !Object.prototype.hasOwnProperty.call(
        ANSWER_KEY[question],
        selectedAnswer
      )
    ) {
      throw new Error("Invalid question or answer");
    }

    return {
      question: question,
      selectedAnswer: selectedAnswer
    };
  });
}

function calculateScores_(responses) {
  // Scores are calculated ONLY from this server-side ANSWER_KEY.
  // Client-supplied score fields are ignored.
  if (!Array.isArray(responses)) {
    throw new Error("Invalid responses");
  }

  let standardScore = 0;
  let firstPreferenceScore = 0;
  let secondPreferenceScore = 0;

  const normalized = normalizeResponses_(responses);

  normalized.forEach(function(r, index) {
    const questionKey = ANSWER_KEY[r.question];
    let points = 0;

    if (
      questionKey &&
      Object.prototype.hasOwnProperty.call(questionKey, r.selectedAnswer)
    ) {
      points = Number(questionKey[r.selectedAnswer]) || 0;
    }

    // The secure HTML sends standard questions first, followed by
    // 4 questions for first preference, then 4 for second preference.
    // We also enforce the expected 11-question application structure.
    if (index < 3) {
      standardScore += points;
    } else if (index < 7) {
      firstPreferenceScore += points;
    } else if (index < 11) {
      secondPreferenceScore += points;
    }
  });

  const totalScore =
    standardScore +
    firstPreferenceScore +
    secondPreferenceScore;

  return {
    standardScore: standardScore,
    firstPreferenceScore: firstPreferenceScore,
    secondPreferenceScore: secondPreferenceScore,
    totalScore: totalScore,
    maxScore: 440
  };
}

// ========================================
// JSON RESPONSE
// ========================================
function json_(obj) {
return ContentService
.createTextOutput(
JSON.stringify(obj)
)
.setMimeType(
ContentService.MimeType.JSON
);
}
