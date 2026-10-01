/**
 * TATTVA BOTANICA — Skin Assessment backend
 * ------------------------------------------
 * WHAT THIS DOES
 * Receives each visitor's submission from index.html (name, phone, age, email,
 * answers, top skin concerns, and 3 selfie photos), writes one row per visitor
 * into this spreadsheet's "Leads" sheet, and saves the 3 photos into a Google
 * Drive folder, storing the photo links in the row. It also handles the
 * "Email me my report" button, sending the visitor their report by email.
 *
 * SETUP (one-time, ~5 minutes)
 * 1. Create a new Google Sheet (e.g. "Tattva Botanica — Skin Assessment Leads").
 * 2. In the Sheet, go to Extensions > Apps Script.
 * 3. Delete any starter code in Code.gs and paste this entire file in its place.
 * 4. Click the disk icon to save. Name the project anything you like.
 * 5. Click "Deploy" > "New deployment".
 *    - Click the gear icon next to "Select type" and choose "Web app".
 *    - Description: "Skin assessment intake" (or anything).
 *    - Execute as: "Me".
 *    - Who has access: "Anyone" (this lets the public HTML page submit to it).
 *    - Click "Deploy". Authorize the permissions it asks for (it's your own script).
 * 6. Copy the "Web app URL" it gives you.
 * 7. Open index.html and paste that URL into the GAS_URL constant near the
 *    top of the <script> section, replacing PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE.
 * 8. Re-deploy this script ("Deploy" > "Manage deployments" > edit > new
 *    version) any time you change this file.
 *
 * The first submission will auto-create:
 *  - A "Leads" sheet tab with headers, if it doesn't already exist.
 *  - A Drive folder named "Tattva Botanica - Skin Assessment Photos"
 *    (created inside the same Drive as this spreadsheet), if it doesn't exist.
 *
 * NOTE ON EMAIL: sending mail uses MailApp under your own Google account, so
 * there's nothing extra to set up — but Google caps free accounts at 100
 * emails/day (1,500/day on Workspace). That's a lot of visitors for one
 * exhibition day; if you expect more, spread deployments across days or ask.
 */

const SHEET_NAME = "Leads";
const PHOTO_FOLDER_NAME = "Tattva Botanica - Skin Assessment Photos";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "email") {
      return handleEmail_(data);
    }
    return handleSubmit_(data);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleSubmit_(data) {
  const sheet = getOrCreateSheet_();
  const folder = getOrCreatePhotoFolder_();

  const visitorFolderName = (data.reportId || data.name || "unknown") + "_" + (data.timestamp || new Date().toISOString());
  const visitorFolder = folder.createFolder(visitorFolderName.replace(/[^a-zA-Z0-9_-]/g, "_"));

  const frontUrl = savePhoto_(visitorFolder, data.photos && data.photos.front, "front");
  const leftUrl  = savePhoto_(visitorFolder, data.photos && data.photos.left, "left");
  const rightUrl = savePhoto_(visitorFolder, data.photos && data.photos.right, "right");

  sheet.appendRow([
    new Date(),                       // A: submitted at (server time)
    data.reportId || "",              // B: Report ID (also shown to visitor)
    data.name || "",                  // C: name
    data.age || "",                   // D: age (optional)
    data.phone || "",                 // E: phone
    data.email || "",                 // F: email (optional)
    data.topConcerns || "",           // G: top skin concerns identified (also shown to visitor)
    data.photoObservations || "",     // H: AI selfie observations (shown to visitor; blank if AI read wasn't usable)
    data.extraNotes || "",            // I: lips/scent/hair/stress notes (also shown to visitor)
    data.suggestedProducts || "",     // J: SUGGESTED PRODUCTS — staff only, never shown to the visitor
    JSON.stringify(data.answers || {}),      // K: raw skin questionnaire answers
    JSON.stringify(data.extraAnswers || {}), // L: raw extra-module answers
    frontUrl,                         // M: front photo link
    leftUrl,                          // N: left photo link
    rightUrl                          // O: right photo link
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleEmail_(data) {
  const to = data.email;
  if (!to || to.indexOf("@") === -1) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "Invalid email" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const subject = "Your Tattva Botanica skin report (" + (data.reportId || "") + ")";
  const body = data.reportText || "Thanks for visiting Tattva Botanica!";
  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: body,
    name: "Tattva Botanica"
  });

  // Optionally log that an email was sent/re-sent, without duplicating the full lead row.
  try {
    const sheet = getOrCreateSheet_();
    sheet.appendRow([
      new Date(), data.reportId || "", data.name || "", "", "", to,
      "(email re-send)", "", "", "", "", "", "", "", ""
    ]);
  } catch (e) { /* logging is best-effort only */ }

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Submitted At", "Report ID", "Name", "Age", "Phone", "Email",
      "Top Skin Concerns", "AI Photo Observations", "Extra Notes (Lips/Scent/Hair/Stress)",
      "Suggested Products (Staff Only)", "Skin Questionnaire Answers (raw)", "Extra Module Answers (raw)",
      "Photo - Front", "Photo - Left", "Photo - Right"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreatePhotoFolder_() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

// Converts a base64 data URL ("data:image/jpeg;base64,....") into a Drive file
// and returns a shareable link, or "" if no photo was provided.
function savePhoto_(folder, dataUrl, label) {
  if (!dataUrl || typeof dataUrl !== "string" || dataUrl.indexOf("base64,") === -1) return "";
  const base64 = dataUrl.split("base64,")[1];
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, "image/jpeg", label + ".jpg");
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // sharing restricted by org policy — file still saved, just not link-shareable
  }
  return file.getUrl();
}
