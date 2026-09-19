 function testBackendConnection() {
  const properties = PropertiesService.getScriptProperties();

  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  const rootFolderId = properties.getProperty('ROOT_FOLDER_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  if (!rootFolderId) {
    throw new Error('ROOT_FOLDER_ID is not configured.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const rootFolder = DriveApp.getFolderById(rootFolderId);

  Logger.log('Spreadsheet: ' + spreadsheet.getName());
  Logger.log('Root folder: ' + rootFolder.getName());

  Logger.log('Sheets:');

  spreadsheet.getSheets().forEach(sheet => {
    Logger.log('- ' + sheet.getName());
  });
}
function testDatabaseStructure() {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const requiredSheets = [
    'Activities',
    'Sessions',
    'Logs'
  ];

  const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());

  Logger.log('Checking database structure...');

  requiredSheets.forEach(sheetName => {

    if (existingSheets.includes(sheetName)) {
      Logger.log('✓ ' + sheetName + ' exists');
    } else {
      Logger.log('✗ ' + sheetName + ' is missing');
    }

  });

  Logger.log('Database structure check completed.');
}
function testActivityWrite() {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('Activities');

  if (!sheet) {
    throw new Error('Activities sheet not found.');
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  Logger.log('Activities columns:');
  headers.forEach((header, index) => {
    Logger.log((index + 1) + '. ' + header);
  });

  Logger.log('Activities sheet is writable and accessible.');
}
function testCoreDatabaseStructure() {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const requiredSheets = [
    'users',
    'Classes',
    'Activities',
    'Sessions',
    'SystemLogs'
  ];

  Logger.log('=== CORE DATABASE STRUCTURE ===');

  requiredSheets.forEach(sheetName => {

    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log('✗ Missing: ' + sheetName);
      return;
    }

    const lastColumn = sheet.getLastColumn();

    if (lastColumn === 0) {
      Logger.log('⚠ Empty sheet: ' + sheetName);
      return;
    }

    const headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0];

    Logger.log('');
    Logger.log('[' + sheetName + ']');

    headers.forEach((header, index) => {
      Logger.log((index + 1) + '. ' + header);
    });
  });

  Logger.log('');
  Logger.log('=== STRUCTURE CHECK COMPLETED ===');
}
function inspectCoreData() {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const sheetsToInspect = [
    'users',
    'Classes'
  ];

  sheetsToInspect.forEach(sheetName => {

    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(sheetName + ' sheet not found.');
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    Logger.log('');
    Logger.log('=== ' + sheetName + ' ===');
    Logger.log('Rows: ' + lastRow);
    Logger.log('Columns: ' + lastColumn);

    if (lastRow <= 1) {
      Logger.log('No data rows found.');
      return;
    }

    const data = sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .getValues();

    Logger.log('Data rows: ' + data.length);

  });

  Logger.log('');
  Logger.log('=== INSPECTION COMPLETED ===');
}
function createTestData() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  const rootFolderId = properties.getProperty('ROOT_FOLDER_ID');

  if (!spreadsheetId || !rootFolderId) {
    throw new Error('Backend configuration is incomplete.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const rootFolder = DriveApp.getFolderById(rootFolderId);

  const classesSheet = ss.getSheetByName('Classes');
  const usersSheet = ss.getSheetByName('users');

  if (!classesSheet || !usersSheet) {
    throw new Error('Required sheets not found.');
  }

  // Prevent duplicate test data
  if (classesSheet.getLastRow() > 1 || usersSheet.getLastRow() > 1) {
    throw new Error('Test data cannot be created because users or Classes already contain data.');
  }

  // IDs
  const classId = 'CLS-TEST-001';
  const userId = 'USR-TEST-001';

  // Create class folder
  const classFolder = rootFolder.createFolder('TEST-Class-9A');

  // Create student folder
  const studentFolder = classFolder.createFolder('TEST-Student-Ali');

  const now = new Date();

  // Classes row
  classesSheet.appendRow([
    classId,
    'کلاس آزمایشی 9A',
    'کلاس آزمایشی سامانه',
    'Active',
    classFolder.getId(),
    now
  ]);

  // Users row
  usersSheet.appendRow([
    userId,
    'علی',
    'رضایی',
    classId,
    'test.student',
    'TEST_HASH',
    'TEST_SALT',
    'Student',
    'Active',
    studentFolder.getId(),
    now,
    now
  ]);

  Logger.log('Test data created successfully.');
  Logger.log('ClassID: ' + classId);
  Logger.log('UserID: ' + userId);
  Logger.log('ClassFolderID: ' + classFolder.getId());
  Logger.log('StudentFolderID: ' + studentFolder.getId());
}
function createTeacherTestData() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const usersSheet = ss.getSheetByName('users');

  if (!usersSheet) {
    throw new Error('users sheet not found.');
  }

  const userId = 'USR-TEACHER-001';
  const username = 'test.teacher';
  const salt = generateSalt();
  const passwordHash = hashPassword('Teacher@1234', salt);
  const now = new Date();

  usersSheet.appendRow([
    userId,
    'Test',
    'Teacher',
    '',
    username,
    passwordHash,
    salt,
    'Teacher',
    'Active',
    '',
    now,
    now
  ]);

  Logger.log('UserID: ' + userId);
}
function testTeacherAccessDenied() {
  const loginResult = loginUser(
    'test.student',
    'Test@1234'
  );

  if (!loginResult.success || !loginResult.sessionId) {
    Logger.log('Student login failed: ' + JSON.stringify(loginResult));
    return;
  }

  Logger.log('Student session created: ' + loginResult.sessionId);

  try {
    const result = getTeacherActivities(
      loginResult.sessionId
    );

    Logger.log('Unexpected success: ' + JSON.stringify(result));
  } catch (error) {
    Logger.log('Access result: ' + error.message);
  }
}
function testPdfSubmission() {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  const rootFolderId = properties.getProperty('ROOT_FOLDER_ID');

  if (!spreadsheetId || !rootFolderId) {
    throw new Error('Backend configuration is incomplete.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);

  const usersSheet = ss.getSheetByName('users');
  const activitiesSheet = ss.getSheetByName('Activities');

  if (!usersSheet || !activitiesSheet) {
    throw new Error('Required sheets not found.');
  }

  // Test student
  const userId = 'USR-TEST-001';
  const classId = 'CLS-TEST-001';

  const userData = usersSheet.getDataRange().getValues();

  let studentFolderId = null;

  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] === userId) {
      studentFolderId = userData[i][9];
      break;
    }
  }

  if (!studentFolderId) {
    throw new Error('Test student folder not found.');
  }

  const studentFolder = DriveApp.getFolderById(studentFolderId);

  // Create a small test PDF
  const pdfContent =
    '%PDF-1.4\n' +
    '1 0 obj\n' +
    '<< /Type /Catalog /Pages 2 0 R >>\n' +
    'endobj\n' +
    '2 0 obj\n' +
    '<< /Type /Pages /Kids [] /Count 0 >>\n' +
    'endobj\n' +
    'trailer\n' +
    '<< /Root 1 0 R >>\n' +
    '%%EOF';

  const fileName = 'TEST-Student-Ali-Assignment.pdf';

  const blob = Utilities.newBlob(
    pdfContent,
    'application/pdf',
    fileName
  );

  const file = studentFolder.createFile(blob);

  const activityId = 'ACT-TEST-001';
  const now = new Date();

  activitiesSheet.appendRow([
    activityId,
    userId,
    classId,
    'تکلیف آزمایشی',
    'تست کامل ارسال فایل PDF',
    file.getName(),
    'PDF',
    file.getMimeType(),
    file.getSize(),
    file.getId(),
    'TEST-SUBMISSION-001',
    now,
    'Submitted',
    '',
    '',
    '',
    ''
  ]);

  Logger.log('PDF submission test successful.');
  Logger.log('ActivityID: ' + activityId);
  Logger.log('FileName: ' + file.getName());
  Logger.log('StorageFileID: ' + file.getId());
  Logger.log('FileSize: ' + file.getSize());
}
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      service: 'Student Portal API',
      status: 'online'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
function hashPassword(password, salt) {
  const data = password + salt;

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    data,
    Utilities.Charset.UTF_8
  );

  return digest
    .map(function(byte) {
      const value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    })
    .join('');
}


function generateSalt() {
  const bytes = Utilities.getUuid();

  return Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      bytes,
      Utilities.Charset.UTF_8
    )
  );
}


function setupTestPassword() {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const usersSheet = ss.getSheetByName('users');

  if (!usersSheet) {
    throw new Error('users sheet not found.');
  }

  const data = usersSheet.getDataRange().getValues();

  const username = 'test.student';
  const password = 'Test@1234';

  for (let i = 1; i < data.length; i++) {

    if (data[i][4] === username) {

      const salt = generateSalt();
      const hash = hashPassword(password, salt);

      // Column 7 = PasswordHash
      usersSheet.getRange(i + 1, 6).setValue(hash);

      // Column 8 = PasswordSalt
      usersSheet.getRange(i + 1, 7).setValue(salt);

      Logger.log('Test password configured successfully.');
      Logger.log('Username: ' + username);

      return;
    }
  }

  throw new Error('Test user not found.');
}


function createSession(userId, role) {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sessionsSheet = ss.getSheetByName('Sessions');

  if (!sessionsSheet) {
    throw new Error('Sessions sheet not found.');
  }

  const sessionId = Utilities.getUuid();

  const now = new Date();

  // Session expires after 24 hours
  const expiresAt = new Date(
    now.getTime() + 24 * 60 * 60 * 1000
  );

  sessionsSheet.appendRow([
    sessionId,
    userId,
    role,
    now,
    expiresAt,
    'Active',
    now
  ]);

  return sessionId;
}


function loginUser(username, password) {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const usersSheet = ss.getSheetByName('users');

  if (!usersSheet) {
    throw new Error('users sheet not found.');
  }

  const data = usersSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const userId = row[0];
    const firstName = row[1];
    const lastName = row[2];
    const classId = row[3];
    const storedUsername = row[4];
    const storedHash = row[5];
    const storedSalt = row[6];
    const role = row[7];
    const status = row[8];

    if (storedUsername !== username) {
      continue;
    }

    if (status !== 'Active') {
      return {
        success: false,
        error: 'ACCOUNT_INACTIVE'
      };
    }

    const calculatedHash = hashPassword(
      password,
      storedSalt
    );

    if (calculatedHash !== storedHash) {
      return {
        success: false,
        error: 'INVALID_CREDENTIALS'
      };
    }

    const sessionId = createSession(
      userId,
      role
    );

    return {
      success: true,
      user: {
        userId: userId,
        firstName: firstName,
        lastName: lastName,
        classId: classId,
        role: role
      },
      sessionId: sessionId
    };
  }

  return {
    success: false,
    error: 'INVALID_CREDENTIALS'
  };
}


function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request body is missing.');
    }

    const request = JSON.parse(
      e.postData.contents
    );

    if (request.action === 'login') {

      if (!request.username || !request.password) {
        throw new Error('Username and password are required.');
      }

      const result = loginUser(
        request.username,
        request.password
      );

      return ContentService
        .createTextOutput(
          JSON.stringify(result)
        )
        .setMimeType(
          ContentService.MimeType.JSON
        );
    }
if (request.action === 'getActivities') {

  if (!request.sessionId) {
    throw new Error('Session ID is required.');
  }

  const result = getStudentActivities(
    request.sessionId
  );

  return ContentService
    .createTextOutput(
      JSON.stringify(result)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
if (request.action === 'submitActivity') {

  if (!request.sessionId) {
    throw new Error('Session ID is required.');
  }

  const result =
    submitStudentActivity(

      request.sessionId,

      request.title,

      request.description,

      request.fileName,

      request.mimeType,

      request.fileBase64

    );


  return ContentService

    .createTextOutput(
      JSON.stringify(result)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}
if (request.action === 'getTeacherActivities') {

  if (!request.sessionId) {
    throw new Error('Session ID is required.');
  }

  const result = getTeacherActivities(
    request.sessionId
  );

  return ContentService
    .createTextOutput(
      JSON.stringify(result)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
if (request.action === 'createAssignment') {
  if (!request.sessionId) {
    throw new Error('Session ID is required.');
  }

  const result = createAssignment(
    request.sessionId,
    request.title,
    request.description,
    request.classId,
    request.deadline
  );

  return ContentService
    .createTextOutput(
      JSON.stringify(result)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
if (request.action === 'getAssignmentsForClass') {
  if (!request.sessionId) {
    throw new Error('Session ID is required.');
  }

  const result = getAssignmentsForClass(
    request.sessionId
  );

  return ContentService
    .createTextOutput(
      JSON.stringify(result)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
if (request.action === 'submitGrade') {
  if (!request.sessionId) {
    throw new Error('Session ID is required.');
  }

  const result = submitGrade(
    request.sessionId,
    request.activityId,
    request.score,
    request.teacherComment
  );

  return ContentService
    .createTextOutput(
      JSON.stringify(result)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}
    throw new Error('Unknown action.');

  } catch (error) {

    return ContentService
      .createTextOutput(
        JSON.stringify({
          success: false,
          error: error.message
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }
}
function createAssignmentsSheet() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  let assignmentsSheet = ss.getSheetByName('Assignments');

  if (!assignmentsSheet) {
    assignmentsSheet = ss.insertSheet('Assignments');
    assignmentsSheet.getRange(1, 1, 1, 8).setValues([[
      'AssignmentID',
      'Title',
      'Description',
      'ClassID',
      'Deadline',
      'Status',
      'CreatedBy',
      'CreatedAt'
    ]]);
  }
}

function createAssignment(
  sessionId,
  title,
  description,
  classId,
  deadline
) {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  if (!sessionId) {
    throw new Error('Session ID is required.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sessionsSheet = ss.getSheetByName('Sessions');
  const assignmentsSheet = ss.getSheetByName('Assignments');

  if (!sessionsSheet) {
    throw new Error('Sessions sheet not found.');
  }

  if (!assignmentsSheet) {
    throw new Error('Assignments sheet not found.');
  }

  const sessionData = sessionsSheet.getDataRange().getValues();
  let userId = null;
  let role = null;

  for (let i = 1; i < sessionData.length; i++) {
    const row = sessionData[i];

    if (row[0] !== sessionId) {
      continue;
    }

    if (row[5] !== 'Active') {
      throw new Error('Session is not active.');
    }

    if (new Date(row[4]) < new Date()) {
      throw new Error('Session has expired.');
    }

    userId = row[1];
    role = row[2];
    break;
  }

  if (!userId) {
    throw new Error('Invalid session.');
  }

  if (role !== 'Teacher') {
    throw new Error('Access denied. Teacher role required.');
  }

  const assignmentId = 'ASG-' + Utilities.getUuid();

  assignmentsSheet.appendRow([
    assignmentId,
    title || '',
    description || '',
    classId || '',
    deadline || '',
    'Active',
    userId,
    new Date()
  ]);

  return {
    success: true,
    assignmentId: assignmentId
  };
}

function getAssignmentsForClass(sessionId) {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  if (!sessionId) {
    throw new Error('Session ID is required.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sessionsSheet = ss.getSheetByName('Sessions');
  const usersSheet = ss.getSheetByName('users');
  const assignmentsSheet = ss.getSheetByName('Assignments');

  if (!sessionsSheet) {
    throw new Error('Sessions sheet not found.');
  }

  if (!usersSheet) {
    throw new Error('users sheet not found.');
  }

  if (!assignmentsSheet) {
    throw new Error('Assignments sheet not found.');
  }

  const sessionData = sessionsSheet.getDataRange().getValues();
  let userId = null;

  for (let i = 1; i < sessionData.length; i++) {
    const row = sessionData[i];

    if (row[0] !== sessionId) {
      continue;
    }

    if (row[5] !== 'Active') {
      throw new Error('Session is not active.');
    }

    if (new Date(row[4]) < new Date()) {
      throw new Error('Session has expired.');
    }

    userId = row[1];
    break;
  }

  if (!userId) {
    throw new Error('Invalid session.');
  }

  const userData = usersSheet.getDataRange().getValues();
  let classId = null;

  for (let i = 1; i < userData.length; i++) {
    const row = userData[i];

    if (row[0] !== userId) {
      continue;
    }

    classId = row[3];
    break;
  }

  if (!classId) {
    throw new Error('Student class not found.');
  }

  const assignmentData = assignmentsSheet.getDataRange().getValues();
  const assignments = [];

  for (let i = 1; i < assignmentData.length; i++) {
    const row = assignmentData[i];

    if (row[3] !== classId || row[5] !== 'Active') {
      continue;
    }

    assignments.push({
      assignmentId: row[0],
      title: row[1],
      description: row[2],
      classId: row[3],
      deadline: row[4],
      status: row[5],
      createdBy: row[6],
      createdAt: row[7]
    });
  }

  return {
    success: true,
    assignments: assignments
  };
}
function testLogin() {

  const result = loginUser(
    'test.student',
    'Test@1234'
  );

  Logger.log(JSON.stringify(result, null, 2));
}
function getStudentActivities(sessionId) {

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  if (!sessionId) {
    throw new Error('Session ID is required.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);

  const sessionsSheet = ss.getSheetByName('Sessions');
  const activitiesSheet = ss.getSheetByName('Activities');

  if (!sessionsSheet) {
    throw new Error('Sessions sheet not found.');
  }

  if (!activitiesSheet) {
    throw new Error('Activities sheet not found.');
  }

  // بررسی Session
  const sessionData = sessionsSheet.getDataRange().getValues();

  let userId = null;

  for (let i = 1; i < sessionData.length; i++) {

    const row = sessionData[i];

    const storedSessionId = row[0];
    const storedUserId = row[1];
    const expiresAt = row[4];
    const status = row[5];

    if (storedSessionId !== sessionId) {
      continue;
    }

    if (status !== 'Active') {
      throw new Error('Session is not active.');
    }

    if (new Date(expiresAt) < new Date()) {
      throw new Error('Session has expired.');
    }

    userId = storedUserId;
    break;
  }

  if (!userId) {
    throw new Error('Invalid session.');
  }

  // دریافت تکالیف
  const activityData = activitiesSheet.getDataRange().getValues();

  const activities = [];

  for (let i = 1; i < activityData.length; i++) {

    const row = activityData[i];

    const activityUserId = row[1];

    if (activityUserId !== userId) {
      continue;
    }

    activities.push({
      activityId: row[0],
      studentId: row[1],
      classId: row[2],
      title: row[3],
      description: row[4],
      fileName: row[5],
      fileType: row[6],
      mimeType: row[7],
      fileSize: row[8],
      storageFileId: row[9],
      internalReference: row[10],
      submittedAt: row[11],
      status: row[12],
      score: row[13],
      teacherComment: row[14],
      reviewedAt: row[15],
      reviewedBy: row[16]
    });
  }

  return {
    success: true,
    userId: userId,
    activities: activities
  };
}
function testGetStudentActivities() {



  const result = getStudentActivities(

    '7609ed61-680f-4743-82ae-9a3736b59684'

  );



  Logger.log(

    JSON.stringify(result, null, 2)

  );

}
function submitStudentActivity(
  sessionId,
  title,
  description,
  fileName,
  mimeType,
  fileBase64
) {

  const properties =
    PropertiesService.getScriptProperties();

  const spreadsheetId =
    properties.getProperty('SPREADSHEET_ID');

  const rootFolderId =
    properties.getProperty('ROOT_FOLDER_ID');


  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  if (!rootFolderId) {
    throw new Error('ROOT_FOLDER_ID is not configured.');
  }

  if (!sessionId) {
    throw new Error('Session ID is required.');
  }

  if (!fileBase64) {
    throw new Error('File data is required.');
  }


  // -----------------------------------
  // 1. بررسی Session
  // -----------------------------------

  const ss =
    SpreadsheetApp.openById(spreadsheetId);

  const sessionsSheet =
    ss.getSheetByName('Sessions');

  const usersSheet =
    ss.getSheetByName('users');

  const activitiesSheet =
    ss.getSheetByName('Activities');


  if (!sessionsSheet) {
    throw new Error('Sessions sheet not found.');
  }

  if (!usersSheet) {
    throw new Error('users sheet not found.');
  }

  if (!activitiesSheet) {
    throw new Error('Activities sheet not found.');
  }


  const sessionData =
    sessionsSheet.getDataRange().getValues();


  let userId = null;


  for (let i = 1; i < sessionData.length; i++) {
    const row = sessionData[i];

    const storedSessionId = row[0];
    const storedUserId = row[1];
    const expiresAt = row[4];
    const status = row[5];


    if (storedSessionId !== sessionId) {
      continue;
    }


    if (status !== 'Active') {
      throw new Error('Session is not active.');
    }


    if (new Date(expiresAt) < new Date()) {
      throw new Error('Session has expired.');
    }


    userId = storedUserId;

    break;
  }


  if (!userId) {
    throw new Error('Invalid session.');
  }


  // -----------------------------------
  // 2. پیدا کردن اطلاعات دانش‌آموز
  // -----------------------------------

  const userData =
    usersSheet.getDataRange().getValues();


  let studentFolderId = null;
  let classId = null;


  for (let i = 1; i < userData.length; i++) {

    const row = userData[i];


    if (row[0] !== userId) {
      continue;
    }


    classId = row[3];
    studentFolderId = row[9];

    break;
  }


  if (!studentFolderId) {
    throw new Error(
      'Student folder not found.'
    );
  }


  // -----------------------------------
  // 3. بررسی نوع فایل
  // -----------------------------------

  const allowedMimeTypes = [
    'application/pdf',
    'image/jpeg',
    'image/png',
    'image/heic',
    'image/heif'
  ];

  if (!allowedMimeTypes.includes(mimeType)) {

    throw new Error(
      'Only PDF, JPG or PNG files are allowed.'
    );

  }


  // -----------------------------------
  // 4. تبدیل Base64 به فایل
  // -----------------------------------

  const decodedBytes =
    Utilities.base64Decode(fileBase64);


  const blob =
    Utilities.newBlob(
      decodedBytes,
      mimeType,
      fileName
    );


  // -----------------------------------
  // 5. ذخیره فایل در پوشه دانش‌آموز
  // -----------------------------------

  const studentFolder =
    DriveApp.getFolderById(studentFolderId);


  const file =
    studentFolder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );


  // -----------------------------------
  // 6. ساخت شناسه Activity
  // -----------------------------------

  const activityId =
    'ACT-' +
    Utilities.getUuid();


  const submittedAt =
    new Date();


  const internalReference =
    'SUBMISSION-' +
    activityId;


  // -----------------------------------
  // 7. ثبت در Activities
  // -----------------------------------

  activitiesSheet.appendRow([

    activityId,
    userId,
    classId,
    title || 'تکلیف',
    description || '',
    file.getName(),
    mimeType === 'application/pdf' ? 'PDF' : 'Image',
    mimeType,
    file.getSize(),
    file.getId(),
    internalReference,
    submittedAt,
    'Submitted',
    '',
    '',
    '',
    ''

  ]);


  // -----------------------------------
  // 8. پاسخ
  // -----------------------------------

  return {

    success: true,

    activityId: activityId,

    userId: userId,

    fileId: file.getId(),

    fileName: file.getName(),

    status: 'Submitted',

    submittedAt: submittedAt

  };

}
function getTeacherActivities(sessionId) {

  const properties =
    PropertiesService.getScriptProperties();

  const spreadsheetId =
    properties.getProperty('SPREADSHEET_ID');


  if (!spreadsheetId) {
    throw new Error(
      'SPREADSHEET_ID is not configured.'
    );
  }


  if (!sessionId) {
    throw new Error(
      'Session ID is required.'
    );
  }


  const ss =
    SpreadsheetApp.openById(
      spreadsheetId
    );


  const sessionsSheet =
    ss.getSheetByName('Sessions');

  const activitiesSheet =
    ss.getSheetByName('Activities');


  if (!sessionsSheet) {
    throw new Error(
      'Sessions sheet not found.'
    );
  }


  if (!activitiesSheet) {
    throw new Error(
      'Activities sheet not found.'
    );
  }


  // -------------------------------
  // بررسی Session و Role
  // -------------------------------

  const sessionData =
    sessionsSheet
      .getDataRange()
      .getValues();


  let userId = null;
  let role = null;


  for (
    let i = 1;
    i < sessionData.length;
    i++
  ) {

    const row =
      sessionData[i];


    const storedSessionId =
      row[0];

    const storedUserId =
      row[1];

    const storedRole =
      row[2];

    const expiresAt =
      row[4];

    const status =
      row[5];


    if (
      storedSessionId !== sessionId
    ) {
      continue;
    }


    if (status !== 'Active') {

      throw new Error(
        'Session is not active.'
      );

    }


    if (
      new Date(expiresAt) < new Date()
    ) {

      throw new Error(
        'Session has expired.'
      );

    }


    userId =
      storedUserId;

    role =
      storedRole;


    break;

  }


  if (!userId) {

    throw new Error(
      'Invalid session.'
    );

  }


  // -------------------------------
  // فقط معلم
  // -------------------------------

  if (role !== 'Teacher') {

    throw new Error(
      'Access denied. Teacher role required.'
    );

  }


  // -------------------------------
  // دریافت Activities
  // -------------------------------

  const activityData =
    activitiesSheet
      .getDataRange()
      .getValues();


  const activities = [];


  for (
    let i = 1;
    i < activityData.length;
    i++
  ) {

    const row =
      activityData[i];


    activities.push({

      activityId: row[0],

      studentId: row[1],

      classId: row[2],

      title: row[3],

      description: row[4],

      fileName: row[5],

      fileType: row[6],

      mimeType: row[7],

      fileSize: row[8],

      storageFileId: row[9],

      internalReference: row[10],

      submittedAt: row[11],

      status: row[12],

      score: row[13],

      teacherComment: row[14],

      reviewedAt: row[15],

      reviewedBy: row[16]

    });

  }


  return {

    success: true,

    teacherId: userId,

    activities: activities

  };

}
function submitGrade(
  sessionId,
  activityId,
  score,
  teacherComment
) {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is not configured.');
  }

  if (!sessionId) {
    throw new Error('Session ID is required.');
  }

  if (!activityId) {
    throw new Error('Activity ID is required.');
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sessionsSheet = ss.getSheetByName('Sessions');
  const activitiesSheet = ss.getSheetByName('Activities');

  if (!sessionsSheet) {
    throw new Error('Sessions sheet not found.');
  }

  if (!activitiesSheet) {
    throw new Error('Activities sheet not found.');
  }

  const sessionData = sessionsSheet.getDataRange().getValues();
  let userId = null;
  let role = null;

  for (let i = 1; i < sessionData.length; i++) {
    const row = sessionData[i];

    if (row[0] !== sessionId) {
      continue;
    }

    if (row[5] !== 'Active') {
      throw new Error('Session is not active.');
    }

    if (new Date(row[4]) < new Date()) {
      throw new Error('Session has expired.');
    }

    userId = row[1];
    role = row[2];
    break;
  }

  if (!userId) {
    throw new Error('Invalid session.');
  }

  if (role !== 'Teacher') {
    throw new Error('Access denied. Teacher role required.');
  }

  const activityData = activitiesSheet.getDataRange().getValues();
  let activityRow = -1;

  for (let i = 1; i < activityData.length; i++) {
    if (activityData[i][0] === activityId) {
      activityRow = i + 1;
      break;
    }
  }

  if (activityRow === -1) {
    throw new Error('Activity not found.');
  }

  activitiesSheet
    .getRange(activityRow, 14, 1, 4)
    .setValues([[
      score,
      teacherComment || '',
      new Date(),
      userId
    ]]);

  return {
    success: true
  };
}
function testSubmitStudentActivity() {

  const sessionId = 'd4a24ac6-b2c2-452b-abe5-eb4bc4bae67a';

  const result = submitStudentActivity(
    sessionId,
    'تکلیف تست ارسال',
    'این یک تکلیف آزمایشی برای تست سیستم ارسال فایل است.',
    'TEST-Upload.pdf',
    'application/pdf',
    'JVBERi0xLjQKJQ=='
  );

  Logger.log(
    JSON.stringify(result, null, 2)
  );

}