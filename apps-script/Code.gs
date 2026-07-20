/**
 * NeoPasture waitlist backend.
 *
 * This file is the SOURCE OF TRUTH. The live copy runs in the Apps Script
 * editor bound to the "Neopasture Waitlist" sheet. To deploy: copy this whole
 * file over the editor contents, then Manage deployments -> pencil -> Version:
 * New version -> Deploy. The web app URL does not change.
 *
 * INVARIANT: customerType is interpolated into buildConfirmationHtml() without
 * escaping. That is safe ONLY because the CUSTOMER_TYPES whitelist check runs
 * first in doPost. Never reorder those, never pass an unvalidated value.
 */

const SHEET_NAME = 'Waitlist';
const CUSTOMER_TYPES = ['Livestock Owner', 'Farm Partner', 'Enterprise Partner'];
const OWNER_EMAIL = 'gbindinazeez@gmail.com';
const RESEND_COOLDOWN_SECONDS = 1800;

const BRAND = {
  ink: '#132A13',
  primary: '#4F772D',
  tint: '#F7FAEB',
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.error('Could not acquire lock; submission rejected.');
    return respond({ ok: false, error: 'Busy right now. Try again in a moment.' });
  }
  try {
    const params = e.parameter;
    if (params['np-check'] || params.company) {
      console.log('Honeypot triggered; submission dropped.');
      return respond({ ok: true });
    }

    const email = String(params.email || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) {
      return respond({ ok: false, error: 'Enter a valid email address.' });
    }

    const phone = String(params.phone || '').trim();
    if (!/^\+?[\d\s().-]{7,20}$/.test(phone) || (phone.match(/\d/g) || []).length < 7) {
      return respond({ ok: false, error: 'Enter a valid phone number.' });
    }

    const customerType = String(params.customerType || '').trim();
    if (CUSTOMER_TYPES.indexOf(customerType) === -1) {
      return respond({ ok: false, error: 'Choose how you want to join.' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date', 'Email', 'Phone', 'Customer Type']);
    }

    if (isDuplicate(sheet, email)) {
      console.log('Duplicate signup, row skipped: ' + email);
    } else {
      sheet.appendRow([new Date(), email, "'" + phone, customerType]);
      SpreadsheetApp.flush();
    }

    const cache = CacheService.getScriptCache();
    const key = cooldownKey(email);
    if (cache.get(key)) {
      console.log('Confirmation suppressed, still in cooldown: ' + email);
    } else if (sendConfirmation(email, customerType)) {
      cache.put(key, '1', RESEND_COOLDOWN_SECONDS);
    }

    return respond({ ok: true });
  } catch (err) {
    console.error('doPost failed: ' + (err && err.stack ? err.stack : err));
    return respond({ ok: false, error: 'Something went wrong. Try again.' });
  } finally {
    lock.releaseLock();
  }
}

function cooldownKey(email) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, email);
  return 'sent:' + Utilities.base64EncodeWebSafe(digest);
}

function isDuplicate(sheet, email) {
  const last = sheet.getLastRow();
  if (last < 2) return false;
  const values = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === email) return true;
  }
  return false;
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Returns true only if Gmail accepted the message. */
function sendConfirmation(email, customerType) {
  try {
    const plain =
      "You're on the waitlist.\n\n" +
      "Thanks for joining the NeoPasture waitlist. We're building NeoPasture now, " +
      "and you'll be among the first to hear when there's a place for you.\n\n" +
      "You're joining as: " + customerType + "\n\n" +
      "No action needed for now. When we're ready to bring you on, we'll reach out " +
      "with your next step. If you have any questions in the meantime, just reply to this email.\n\n" +
      "NeoPasture\n" +
      "hello@neopasture.com\n" +
      "https://neopasture.com";

    GmailApp.sendEmail(email, "You're on the NeoPasture waitlist", plain, {
      from: 'olaitan@neopasture.com',
      name: 'Olaitan from NeoPasture',
      replyTo: 'hello@neopasture.com',
      htmlBody: buildConfirmationHtml(customerType),
    });
    return true;
  } catch (err) {
    console.error('Confirmation email failed: ' + (err && err.stack ? err.stack : err));
    return false;
  }
}

function buildConfirmationHtml(customerType) {
  return '' +
'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" bgcolor="#EEF2E0" style="background-color:#EEF2E0;">' +
  '<tr><td align="center" style="padding:32px 16px;">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:600px;max-width:600px;background-color:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #DDE4C8;">' +
      '<tr><td align="center" bgcolor="#4F772D" style="background-color:#4F772D;padding:34px 40px;">' +
        '<div style="font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:22px;font-weight:700;letter-spacing:0.5px;color:#ffffff;">NeoPasture</div>' +
        '<div style="font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:12px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;color:#ECF39E;padding-top:6px;">Livestock, digitised</div>' +
      '</td></tr>' +
      '<tr><td style="padding:40px 40px 8px 40px;">' +
        '<h1 style="margin:0 0 16px 0;font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:24px;line-height:1.3;font-weight:700;color:#132A13;">You\'re on the waitlist.</h1>' +
        '<p style="margin:0 0 20px 0;font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:16px;line-height:1.65;color:#3A4A33;">Thanks for joining the NeoPasture waitlist. We\'re building NeoPasture now, and you\'ll be among the first to hear when there\'s a place for you.</p>' +
      '</td></tr>' +
      '<tr><td style="padding:0 40px 24px 40px;">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" bgcolor="#ECF39E" style="background-color:#ECF39E;border-radius:10px;">' +
          '<tr><td style="padding:16px 20px;">' +
            '<div style="font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:11px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;color:#4F772D;">You\'re joining as</div>' +
            '<div style="font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:18px;font-weight:700;color:#132A13;padding-top:4px;">' + customerType + '</div>' +
          '</td></tr>' +
        '</table>' +
      '</td></tr>' +
      '<tr><td style="padding:0 40px 8px 40px;">' +
        '<p style="margin:0 0 20px 0;font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:16px;line-height:1.65;color:#3A4A33;">No action needed for now. When we\'re ready to bring you on, we\'ll reach out with your next step. If you have any questions in the meantime, just reply to this email.</p>' +
      '</td></tr>' +
      '<tr><td style="padding:8px 40px 36px 40px;">' +
        '<table role="presentation" cellpadding="0" cellspacing="0"><tr><td align="center" bgcolor="#4F772D" style="background-color:#4F772D;border-radius:8px;">' +
          '<a href="https://neopasture.com" style="display:inline-block;padding:14px 28px;font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:15px;font-weight:600;color:#ffffff;text-decoration:none;">Explore NeoPasture</a>' +
        '</td></tr></table>' +
      '</td></tr>' +
      '<tr><td style="padding:0 40px;"><div style="height:2px;background-color:#90A955;font-size:0;line-height:0;">&nbsp;</div></td></tr>' +
      '<tr><td style="padding:24px 40px 34px 40px;">' +
        '<p style="margin:0 0 6px 0;font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:13px;line-height:1.6;color:#6B7A5E;">NeoPasture &middot; <a href="mailto:hello@neopasture.com" style="color:#4F772D;text-decoration:none;">hello@neopasture.com</a></p>' +
        '<p style="margin:0;font-family:\'Space Grotesk\',Helvetica,Arial,sans-serif;font-size:12px;line-height:1.6;color:#9AA88A;">You\'re receiving this because you joined the waitlist at neopasture.com.</p>' +
      '</td></tr>' +
    '</table>' +
  '</td></tr>' +
'</table>';
}

/** Run manually from the editor. Applies brand styling to the sheet. */
function formatWaitlistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Email', 'Phone', 'Customer Type']);
  }

  const maxRows = sheet.getMaxRows();

  const bandings = sheet.getBandings();
  for (let i = 0; i < bandings.length; i++) {
    bandings[i].remove();
  }
  sheet.getRange(1, 1, maxRows, 4)
    .applyRowBanding()
    .setHeaderRowColor(BRAND.primary)
    .setFirstRowColor('#FFFFFF')
    .setSecondRowColor(BRAND.tint)
    .setFooterRowColor(null);

  sheet.getRange(1, 1, 1, 4)
    .setFontFamily('Space Grotesk')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 40);

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 180);

  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, 4)
      .setFontFamily('Space Grotesk')
      .setFontColor(BRAND.ink)
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat('yyyy-mm-dd  hh:mm');
  }

  console.log('Waitlist sheet formatted.');
}

/** Run manually from the editor after any change. Grants scopes and previews. */
function authorizeAndTest() {
  console.log('Gmail aliases: ' + GmailApp.getAliases().join(', '));
  console.log('Mail quota left today: ' + MailApp.getRemainingDailyQuota());
  sendConfirmation(OWNER_EMAIL, 'Livestock Owner');
  console.log('Preview confirmation sent to ' + OWNER_EMAIL);
}
