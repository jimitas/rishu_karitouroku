// データアクセス層: スプレッドシートとのデータ操作を担当

/**
 * 現在ログインしているユーザーのメールアドレスを取得
 */
function getCurrentUserEmail() {
  try {
    const user = Session.getActiveUser();
    const email = user.getEmail();

    if (!email) {
      throw new Error('ユーザー情報を取得できませんでした。');
    }

    return {
      success: true,
      email: email,
      message: 'ユーザー認証成功'
    };
  } catch (error) {
    return {
      success: false,
      email: null,
      message: 'ユーザー認証に失敗しました: ' + error.message
    };
  }
}

/**
 * メールアドレスから学生情報を取得
 */
function getStudentInfoByEmail(email) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修情報');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) { // A列（メールアドレス）と照合
      return {
        success: true,
        student: {
          email: data[i][0],     // メールアドレス
          id: data[i][1],        // 学籍番号
          name: data[i][5],      // 氏名
          grade: data[i][2],     // 学年
          class: data[i][3],     // クラス
          number: data[i][4]     // 番号
        },
        message: '学生情報が見つかりました。'
      };
    }
  }

  return {
    success: false,
    student: null,
    message: 'このメールアドレスは履修情報に登録されていません。管理者にお問い合わせください。'
  };
}

/**
 * 学生情報を取得（履修情報シートから）
 */
function getStudentInfo(studentId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修情報');
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // ヘッダー行を取得

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) { // 学籍番号は2列目（インデックス1）
      return {
        id: data[i][1],        // 学籍番号
        email: data[i][0],     // メールアドレス
        name: data[i][5],      // 氏名
        grade: data[i][2],     // 学年
        class: data[i][3],     // クラス
        number: data[i][4]     // 番号
      };
    }
  }

  return null;
}

/**
 * 科目情報を取得（履修情報シートのヘッダーから生成）
 */
function getSubjects() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修情報');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const subjects = [];

  // 7列目以降が科目列（0ベースで6以降）
  for (let col = 6; col < headers.length; col++) {
    const subjectName = headers[col];

    if (subjectName && subjectName.trim() !== '') {
      // 科目名から学年を推測（Ⅰ、Ⅱ、Ⅲの有無で判定）
      let targetGrade = '1'; // デフォルトは1年
      if (subjectName.includes('Ⅱ')) {
        targetGrade = '2';
      } else if (subjectName.includes('Ⅲ')) {
        targetGrade = '3';
      }

      // 科目の種別を推測（基本的な分類）
      let type = '選択';
      let subject = '一般';

      // 必修科目の判定（一般的な必修科目名で判定）
      const requiredSubjects = ['現代の国語', '歴史総合', '地理総合', '数学Ⅰ', '科学と人間生活', '体育Ⅰ', '英語コミュニケーションⅠ', '情報Ⅰ'];
      if (requiredSubjects.some(req => subjectName.includes(req.replace(/ｺﾐｭﾆｹｰｼｮﾝ/g, 'コミュニケーション')))) {
        type = '必修';
      }

      // 教科分類
      if (subjectName.includes('国語') || subjectName.includes('文学') || subjectName.includes('言語文化')) {
        subject = '国語';
      } else if (subjectName.includes('歴史') || subjectName.includes('地理') || subjectName.includes('公共')) {
        subject = '地理歴史・公民';
      } else if (subjectName.includes('数学')) {
        subject = '数学';
      } else if (subjectName.includes('化学') || subjectName.includes('生物') || subjectName.includes('科学')) {
        subject = '理科';
      } else if (subjectName.includes('体育') || subjectName.includes('保健')) {
        subject = '保健体育';
      } else if (subjectName.includes('美術')) {
        subject = '芸術';
      } else if (subjectName.includes('英語') || subjectName.includes('論理')) {
        subject = '外国語';
      } else if (subjectName.includes('家庭')) {
        subject = '家庭';
      } else if (subjectName.includes('情報')) {
        subject = '情報';
      } else if (subjectName.includes('総合') || subjectName.includes('創造')) {
        subject = '総合・専門';
      }

      subjects.push({
        id: `SUBJ_${col}`,     // 列番号ベースのID
        name: subjectName,     // 科目名
        type: type,           // 必修/選択
        subject: subject,     // 教科分類
        credits: 2,          // デフォルト単位数（後で設定可能にする）
        targetGrade: targetGrade, // 対象学年
        available: true      // 履修可能
      });
    }
  }

  return subjects;
}

/**
 * 学年別の履修可能科目を取得
 */
function getAvailableSubjects(grade, studentId) {
  const allSubjects = getSubjects();
  const studentRegistrations = getStudentRegistrations(studentId);
  const registeredSubjectNames = studentRegistrations.map(reg => reg.subjectName);

  // 学年に応じた履修可能科目をフィルタリング
  const availableSubjects = allSubjects.filter(subject => {
    // 既に履修登録済みの科目は除外
    if (registeredSubjectNames.includes(subject.name)) {
      return false;
    }

    // 対象学年以下の科目のみ表示（例: 2年生は1年・2年科目を履修可能）
    return parseInt(subject.targetGrade) <= parseInt(grade);
  });

  return availableSubjects;
}

/**
 * 学生の履修登録情報を取得（履修情報シートから）
 */
function getStudentRegistrations(studentId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修情報');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const registrations = [];

  // 学生データを見つける
  let studentRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) {
      studentRowIndex = i;
      break;
    }
  }

  if (studentRowIndex === -1) {
    return [];
  }

  // 科目列（7列目以降）をチェックして履修済みの科目を抽出
  for (let col = 6; col < headers.length; col++) {
    const subjectName = headers[col];
    const registrationStatus = data[studentRowIndex][col];

    // 値があれば履修登録済みとみなす
    if (registrationStatus && registrationStatus !== '') {
      registrations.push({
        subjectName: subjectName,
        status: registrationStatus,
        studentId: studentId
      });
    }
  }

  return registrations;
}

/**
 * 履修登録（履修情報シートの該当科目列を更新）
 */
function registerSubject(studentId, subjectName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修情報');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 学生の行を見つける
  let studentRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) {
      studentRowIndex = i;
      break;
    }
  }

  if (studentRowIndex === -1) {
    return { success: false, message: '学生が見つかりませんでした。' };
  }

  // 科目名の列を見つける
  let subjectColIndex = -1;
  for (let col = 6; col < headers.length; col++) {
    if (headers[col] === subjectName) {
      subjectColIndex = col;
      break;
    }
  }

  if (subjectColIndex === -1) {
    return { success: false, message: '科目が見つかりませんでした。' };
  }

  // 既に登録済みかチェック
  if (data[studentRowIndex][subjectColIndex] && data[studentRowIndex][subjectColIndex] !== '') {
    return { success: false, message: 'この科目は既に履修登録済みです。' };
  }

  // 履修登録（セルに「履修中」を設定）
  sheet.getRange(studentRowIndex + 1, subjectColIndex + 1).setValue('履修中');

  return { success: true, message: '履修登録が完了しました。' };
}

/**
 * 履修取消（履修情報シートの該当科目列をクリア）
 */
function unregisterSubject(studentId, subjectName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('履修情報');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 学生の行を見つける
  let studentRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === studentId) {
      studentRowIndex = i;
      break;
    }
  }

  if (studentRowIndex === -1) {
    return { success: false, message: '学生が見つかりませんでした。' };
  }

  // 科目名の列を見つける
  let subjectColIndex = -1;
  for (let col = 6; col < headers.length; col++) {
    if (headers[col] === subjectName) {
      subjectColIndex = col;
      break;
    }
  }

  if (subjectColIndex === -1) {
    return { success: false, message: '科目が見つかりませんでした。' };
  }

  // 履修登録されていなければエラー
  if (!data[studentRowIndex][subjectColIndex] || data[studentRowIndex][subjectColIndex] === '') {
    return { success: false, message: 'この科目は履修登録されていません。' };
  }

  // 履修取消（セルをクリア）
  sheet.getRange(studentRowIndex + 1, subjectColIndex + 1).setValue('');

  return { success: true, message: '履修取消が完了しました。' };
}

