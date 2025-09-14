// メイン機能: Webアプリケーション

/**
 * WebアプリのメインHTML表示
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('履修登録システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLテンプレートでファイルをインクルードする関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * 学年に応じた履修可能科目を取得
 */
function getAvailableSubjects(studentGrade, studentId) {
  const subjects = getSubjects();
  const registrations = getStudentRegistrations(studentId);
  const registeredSubjectIds = registrations.map(reg => reg.subjectId);

  return subjects.filter(subject => {
    // 既に履修済みの科目は除外
    if (registeredSubjectIds.includes(subject.id)) {
      return false;
    }

    // 学年制限チェック（学年条件列を参照）
    if (subject.gradeCondition && parseInt(subject.gradeCondition) > parseInt(studentGrade)) {
      return false;
    }

    // 履修条件チェック（科目情報内のprerequisite列を参照）
    if (subject.prerequisite && subject.prerequisite !== '' && subject.prerequisite !== 'なし') {
      // TODO: より詳細な前提科目チェックを実装
      // 現在は簡単な文字列チェックのみ
    }

    return true;
  });
}

