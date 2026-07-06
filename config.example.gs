/**
 * 個人情報設定のテンプレート（プレースホルダ）。このファイルはリポジトリにコミットする。
 *
 * セットアップ手順:
 *  1. このファイルを config.local.gs という名前でコピーする（config.local.gs は .gitignore 済み）。
 *  2. config.local.gs の各値を実際の値に書き換える。
 *  3. Apps Script エディタで installAppConfig を一度だけ実行する。
 *     → Script Properties `APP_CONFIG` に JSON が保存され、Code.gs の cfg_() が読み込む。
 *  4. 値を変更したいときは config.local.gs を直して再度 installAppConfig を実行する。
 *
 * 注意: 実際の個人情報は config.local.gs（非追跡）にのみ置き、このファイルには書かない。
 */
function installAppConfig() {
  const config = {
    allowedEmail: 'sender@example.com',   // このアプリを使える唯一のGoogleアカウント
    userId: 'user_example',                // 内部データの所有者ID（任意の固定文字列）
    userName: '氏名',                      // 送信者名（メール本文・件名に使用）
    recipient: 'recipient@example.com',    // 宛先(To)
    recipientName: 'ご担当者',             // メール本文の宛名（敬称込み）
    ccRecipients: ['cc1@example.com', 'cc2@example.com'],
    commuteRoute: '自宅ー勤務先',          // メール本文の交通費・経路
    commuteCost: '0円(定期)',
    salesFileName: '作業報告書.xlsx',      // 営業用ファイル名
    clientFilePerson: '氏名',              // 客先用ファイル名の氏名部分（出勤簿_YYYY年MM月_◯◯.xlsx）
    // 営業用テンプレは失われた「入力用」ブックへの外部参照を持ち、Sheets変換で #ERROR になる。
    // その代替として下記の実値をコードから記入する。
    salesReport: {
      managementNumber: '管理番号',        // M1
      employeeName: '氏名　フルネーム',    // C6（全角スペース区切り）
      closingLabel: '月末締め',            // N5
      calcUnit: 10,                        // G5/H5
      contract: {
        D48: '派遣先名称',
        M48: '派遣元名称',
        D49: '事業所',
        D50: '所在地',
        D51: '組織単位',
        D52: '業務の種類',
        D53: '業務内容',
        D56: '責任の程度',
        D57: '就業場所①',
        D58: '就業場所②',
        D59: '就業場所③',
        Q66: 'フッター',
      },
      shift: { start: '09:00', end: '17:40', breakStart: '12:00', breakEnd: '13:00' },
    },
  };
  PropertiesService.getScriptProperties().setProperty('APP_CONFIG', JSON.stringify(config));
  Logger.log('APP_CONFIG を保存しました。');
}
