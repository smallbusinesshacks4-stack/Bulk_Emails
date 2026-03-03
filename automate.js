/***** CONFIG *****/
const SHEET_NAME = "Form Responses 1";  // change if your tab has a different name
const DRY_RUN = true;                  // keep true first to test; then set false to send
const MAX_SEND_PER_RUN = 500;           // safety cap
const DELAY_MS = 1200;                  // delay between emails

// These MUST match your sheet headers exactly:
const FIRST_NAME_HEADER = "First Name";
const EMAIL_HEADER = "Email address";
const SENT_HEADER = "sent";
const SENT_AT_HEADER = "sent_at";

// Your email content
const SUBJECT_TEMPLATE =
  "Congratulations {{First Name}}, Your discount proposal has been accepted";

const BODY_TEMPLATE = `
Hi {{First Name}},

Thanks for signing up with us. Our product is the best and we are happy to give you 30% discount.

Regards

Tracy,
Marketing Lead
`.trim();

/***** STATUS LOGGER *****/
function logSendStatus() {
  const remaining = MailApp.getRemainingDailyQuota();
  const lastRun = PropertiesService.getScriptProperties().getProperty("LAST_SEND_TIME");

  Logger.log("=== EMAIL STATUS ===");
  Logger.log("Remaining daily email quota: " + remaining);

  if (lastRun) {
    const lastDate = new Date(lastRun);
    const nextPossible = new Date(lastDate.getTime() + 24 * 60 * 60 * 1000);
    Logger.log("Last email sent at: " + lastDate);
    Logger.log("Estimated quota reset at: " + nextPossible);
  } else {
    Logger.log("No previous send recorded.");
  }

  Logger.log("====================");
}

/***** MAIN *****/
function sendPersonalizedEmails() {
  logSendStatus();

  // Guard: stop early if quota is exhausted (prevents your “invoked too many times” crash)
  if (!DRY_RUN) {
    const remaining = MailApp.getRemainingDailyQuota();
    if (remaining <= 0) {
      Logger.log("Daily GmailApp quota exhausted. Stopping run.");
      return;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("No data rows found.");

  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  const firstNameCol = headers.indexOf(FIRST_NAME_HEADER);
  const emailCol = headers.indexOf(EMAIL_HEADER);
  const sentCol = headers.indexOf(SENT_HEADER);
  const sentAtCol = headers.indexOf(SENT_AT_HEADER);

  if (firstNameCol === -1) throw new Error(`Missing header: ${FIRST_NAME_HEADER}`);
  if (emailCol === -1) throw new Error(`Missing header: ${EMAIL_HEADER}`);
  if (sentCol === -1) throw new Error(`Missing header: ${SENT_HEADER} (please add this column)`);
  if (sentAtCol === -1) throw new Error(`Missing header: ${SENT_AT_HEADER} (please add this column)`);

  // Cap per run by remaining quota (so MAX_SEND_PER_RUN can’t exceed what’s left)
  let cap = MAX_SEND_PER_RUN;
  if (!DRY_RUN) cap = Math.min(cap, MailApp.getRemainingDailyQuota());

  let processed = 0;

  for (let r = 0; r < rows.length; r++) {
    if (processed >= cap) break;

    const sheetRowNumber = r + 2; // data starts at row 2
    const row = rows[r];

    const alreadySent = String(row[sentCol] || "").toLowerCase().trim();
    if (["sent", "yes", "true", "1"].includes(alreadySent)) continue;

    const to = String(row[emailCol] || "").trim();
    if (!to) continue;

    // Build data object keyed by the EXACT header names
    const data = {};
    headers.forEach((h, i) => (data[h] = row[i]));

    const subject = applyTemplate(SUBJECT_TEMPLATE, data);
    const body = applyTemplate(BODY_TEMPLATE, data).replace(/\n{3,}/g, "\n\n");

    if (DRY_RUN) {
      Logger.log(`[DRY_RUN] Would send to: ${to} | Subject: ${subject}`);
    } else {
      GmailApp.sendEmail(to, subject, body);

      // Record last send time (this was missing in your version)
      PropertiesService.getScriptProperties().setProperty(
        "LAST_SEND_TIME",
        new Date().toISOString()
      );

      // Mark sheet
      sheet.getRange(sheetRowNumber, sentCol + 1).setValue("SENT");
      sheet.getRange(sheetRowNumber, sentAtCol + 1).setValue(new Date());

      Utilities.sleep(DELAY_MS);
    }

    processed++;
  }

  Logger.log(`Done. Processed: ${processed}`);
}

/***** HELPER *****/
function applyTemplate(template, data) {
  return template.replace(/{{\s*([^}]+)\s*}}/g, (_, key) => {
    const k = String(key).trim();
    const val = data[k];
    return val === null || val === undefined ? "" : String(val);
  });
}
