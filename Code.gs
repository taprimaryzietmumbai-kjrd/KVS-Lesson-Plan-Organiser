const CONFIG = {
  SPREADSHEET_ID: '1yauP7QTxR3CyGrU9lKhuWOWnrvdDbQCszrTrKnabnLM',
  SHEET_NAME: 'LessonPlans',
  PARTICIPANT_SPREADSHEET_ID: '1pksDgeV53waKMP1pmj0CWbpP3qyAT8nrN8RhU9Y-mKE',
  PARTICIPANT_SHEET_NAME: 'PRTM ADI',
  ADMIN_EMAIL: 'taprimayzietmumbai@gmail.com',
  OTP_TTL_SECONDS: 300,
  OTP_RESEND_SECONDS: 60,
  OTP_MAX_ATTEMPTS: 5,
  ADMIN_SESSION_SECONDS: 1800
};

const HEADERS = ['ID','Updated At','Teacher Name','Designation','Class','Subject','Chapter','Date','Periods Required','Students','Source','Developer','Complete Lesson Plan JSON'];

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const params = (e && e.parameter) || {};
    let data = params;
    if (e && e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
    const action = String(params.action || data.action || 'ping');

    if (action === 'ping') return output({success:true,message:'KVS Lesson Plan Organiser connected.'});
    if (action === 'participants') return output({success:true,participants:getParticipants_()});
    if (action === 'sendAdminOtp') return output(sendAdminOtp_());
    if (action === 'verifyAdminOtp') return output(verifyAdminOtp_(String(data.otp || '')));

    // Saved-plan listing, retrieval and deletion are authorised on the server.
    if (['list','get','delete'].indexOf(action) !== -1) {
      const auth = requireAdmin_(String(data.adminToken || params.adminToken || ''));
      if (!auth.ok) return output({success:false,message:auth.message,code:'ADMIN_AUTH_REQUIRED'});
    }

    const sheet = getSheet_();
    if (action === 'save') {
      const id = data.id || Utilities.getUuid();
      const values = sheet.getDataRange().getValues();
      let rowNumber = 0;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(id)) { rowNumber = i + 1; break; }
      }
      const record = [
        id,new Date(),data.teacherName || '',data.designation || '',
        data.className || '',data.subject || '',data.chapter || '',
        data.date || '',data.periodsRequired || '',data.students || '',
        data.source || '',data.developer || '',JSON.stringify(data)
      ];
      if (rowNumber) sheet.getRange(rowNumber,1,1,record.length).setValues([record]);
      else sheet.appendRow(record);
      return output({success:true,id:id,message:'Lesson plan saved successfully.'});
    }

    if (action === 'list') {
      const values = sheet.getDataRange().getValues();
      const plans = [];
      for (let i = 1; i < values.length; i++) plans.push({
        id:values[i][0],updatedAt:values[i][1],teacherName:values[i][2],
        designation:values[i][3],className:values[i][4],subject:values[i][5],
        chapter:values[i][6],date:values[i][7]
      });
      return output({success:true,plans:plans});
    }

    if (action === 'get') {
      const id = data.id || params.id;
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(id)) return output({success:true,plan:JSON.parse(values[i][12])});
      }
      return output({success:false,message:'Lesson plan not found.'});
    }

    if (action === 'delete') {
      const id = data.id || params.id;
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(id)) {
          sheet.deleteRow(i + 1);
          return output({success:true,message:'Lesson plan deleted successfully.'});
        }
      }
      return output({success:false,message:'Lesson plan not found.'});
    }
    return output({success:false,message:'Unknown action: ' + action});
  } catch (err) {
    return output({success:false,message:err && err.message ? err.message : String(err)});
  }
}

// Return only the six approved profile fields required by the lesson-plan form.
function getParticipants_() {
  const ss = SpreadsheetApp.openById(CONFIG.PARTICIPANT_SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.PARTICIPANT_SHEET_NAME);
  if (!sh) throw new Error('Participant tab "' + CONFIG.PARTICIPANT_SHEET_NAME + '" was not found.');
  const rows = sh.getDataRange().getDisplayValues();
  if (rows.length < 2) return [];

  const headers = rows[0].map(v => String(v || '').trim().toLowerCase().replace(/\\s+/g,' '));
  function findHeader(names) {
    for (let i = 0; i < names.length; i++) {
      const index = headers.indexOf(names[i].toLowerCase());
      if (index !== -1) return index;
    }
    return -1;
  }
  const cols = {
    name:findHeader(['name of the participant']),
    designation:findHeader(['designation']),
    school:findHeader(['school']),
    region:findHeader(['region']),
    employeeCode:findHeader(['employee code']),
    joiningYear:findHeader(['year of joining in kvs']),
    educationalQualification:findHeader(['edu. qlfn.','edu qlfn.','educational qualification'])
  };
  if (cols.name === -1) throw new Error('Name of the Participant column not found in PRTM ADI.');

  const seen = {};
  const participants = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const name = String(row[cols.name] || '').trim();
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen[key]) continue;
    seen[key] = true;
    participants.push({
      name:name,
      designation:cols.designation >= 0 ? String(row[cols.designation] || '').trim() : '',
      school:cols.school >= 0 ? String(row[cols.school] || '').trim() : '',
      region:cols.region >= 0 ? String(row[cols.region] || '').trim() : '',
      employeeCode:cols.employeeCode >= 0 ? String(row[cols.employeeCode] || '').trim() : '',
      joiningYear:cols.joiningYear >= 0 ? String(row[cols.joiningYear] || '').trim() : '',
      educationalQualification:cols.educationalQualification >= 0 ? String(row[cols.educationalQualification] || '').trim() : ''
    });
  }
  return participants;
}

function sendAdminOtp_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('admin_otp_throttle')) return {success:false,message:'Please wait 60 seconds before requesting another OTP.'};
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const salt = Utilities.getUuid();
  cache.put('admin_otp_record',JSON.stringify({salt:salt,hash:sha256_(salt + ':' + otp),attempts:0}),CONFIG.OTP_TTL_SECONDS);
  cache.put('admin_otp_throttle','1',CONFIG.OTP_RESEND_SECONDS);
  MailApp.sendEmail({
    to:CONFIG.ADMIN_EMAIL,
    subject:'KVS Lesson Plan Organiser — Administrator OTP',
    body:'Your administrator verification code is ' + otp + '. It expires in 5 minutes. Do not share it.'
  });
  return {success:true,message:'OTP sent to the administrator email address. It is valid for 5 minutes.'};
}

function verifyAdminOtp_(otp) {
  if (!/^\\d{6}$/.test(otp)) return {success:false,message:'Please enter the 6-digit OTP.'};
  const cache = CacheService.getScriptCache();
  const raw = cache.get('admin_otp_record');
  if (!raw) return {success:false,message:'OTP expired. Please request a new OTP.'};
  const record = JSON.parse(raw);
  if (record.attempts >= CONFIG.OTP_MAX_ATTEMPTS) {
    cache.remove('admin_otp_record');
    return {success:false,message:'Maximum attempts reached. Request a new OTP.'};
  }
  if (sha256_(record.salt + ':' + otp) !== record.hash) {
    record.attempts++;
    cache.put('admin_otp_record',JSON.stringify(record),CONFIG.OTP_TTL_SECONDS);
    return {success:false,message:'Incorrect OTP. Attempts remaining: ' + (CONFIG.OTP_MAX_ATTEMPTS - record.attempts)};
  }
  cache.remove('admin_otp_record');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('admin_session_' + sha256_(token),'valid',CONFIG.ADMIN_SESSION_SECONDS);
  return {success:true,adminToken:token,expiresIn:CONFIG.ADMIN_SESSION_SECONDS,message:'Administrator access verified.'};
}

function requireAdmin_(token) {
  if (!token || token.length < 30) return {ok:false,message:'Administrator OTP verification is required.'};
  return CacheService.getScriptCache().get('admin_session_' + sha256_(token)) === 'valid'
    ? {ok:true} : {ok:false,message:'Administrator session expired. Please verify by OTP again.'};
}

function sha256_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,value,Utilities.Charset.UTF_8);
  return bytes.map(b => {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  return sh;
}

function setup() {
  const sh = getSheet_();
  SpreadsheetApp.flush();
  Logger.log('Setup completed. Sheet: ' + sh.getName() + '; Spreadsheet ID: ' + CONFIG.SPREADSHEET_ID);
}

function output(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
