// データアクセス層: スプレッドシートとのデータ操作を担当

/**
 * 学生情報を取得
 */
function getStudentInfo(studentId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('学生情報');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      return {
        id: data[i][0],
        name: data[i][1],
        grade: data[i][2],
        class: data[i][3],
        enrollmentYear: data[i][4]
      };
    }
  }

  return null;
}

/**
 * 科目情報を取得
 */
function getSubjects() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('科目情報');
  const data = sheet.getDataRange().getValues();
  const subjects = [];

  for (let i = 1; i < data.length; i++) {
    subjects.push({
      id: data[i][0],        // 科目コード
      name: data[i][1],      // 科目名
      type: data[i][2],      // 種別
      subject: data[i][3],   // 教科名
      credits: data[i][4],   // 単位数
      required: data[i][5],  // 必修
      gradeCondition: data[i][6],  // 学年条件
      prerequisite: data[i][7],    // 履修条件
      textbook: data[i][8],        // 教科書
      supplementary: data[i][9]    // 副教材
    });
  }

  return subjects;
}

/**
 * 学生の履修登録情報を取得
 */
function getStudentRegistrations(studentId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修登録');
  const data = sheet.getDataRange().getValues();
  const registrations = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) {
      registrations.push({
        registrationId: data[i][0],
        studentId: data[i][1],
        subjectId: data[i][2],
        year: data[i][3],
        registrationDate: data[i][4],
        status: data[i][5]
      });
    }
  }

  return registrations;
}

/**
 * 履修登録
 */
function registerSubject(studentId, subjectId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修登録');

  // 新しい登録IDを生成
  const data = sheet.getDataRange().getValues();
  const newId = 'REG' + String(data.length).padStart(3, '0');

  // 現在の年度を取得
  const currentYear = new Date().getFullYear();

  // 新しい行を追加
  sheet.appendRow([
    newId,
    studentId,
    subjectId,
    currentYear,
    new Date(),
    '登録済み'
  ]);

  return { success: true, message: '履修登録が完了しました。' };
}

/**
 * 履修取消
 */
function unregisterSubject(studentId, subjectId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修登録');
  const data = sheet.getDataRange().getValues();

  // 該当する履修登録を探して削除
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === studentId && data[i][2] === subjectId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: '履修取消が完了しました。' };
    }
  }

  return { success: false, message: '履修登録が見つかりませんでした。' };
}

