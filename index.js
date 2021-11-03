const fs = require('fs')
const XLSX = require('xlsx')

const fileName = "ExcelToJson"; // ファイル名を入れる（この名前が JSON ファイル名になる）
const workbook = XLSX.readFile(`excel/${fileName}.xlsx`)

// 1．データ取得(JSON)
let bookCategory, bookProfileDirector, bookProfileGroup

workbook.SheetNames.forEach(sheet => {
  // 各シート名の変数名を指定
  if ("項目名" === sheet) bookCategory = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
  if ("役職" === sheet) bookProfileDirector = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
  if ("グループ" === sheet) bookProfileGroup = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
})

const bookProfiles = bookProfileDirector.concat(bookProfileGroup) //各 Profile データを結合

// 2. データの一部を変換 (key:日本語表記 => 英語表記 / 改行文字 => HTML タグ)
let keyEN

const newBookProfiles = bookProfiles.map(profileData => {
  const newProfileData = {}
  for (const key in profileData) {
    if (Object.hasOwnProperty.call(profileData, key)) {
      replaceKeyName(key) // key の英語表記を取得
      newProfileData[keyEN] = profileData[key].replace(/[\n]/g, '<br>') // 改行文字を HTML タグに変換し代入
    }
  }
  return newProfileData
})

// 項目名シートに倣って、日本語表記の key を英語表記で返す
function replaceKeyName(keyJP) {
  bookCategory.map(categoryData => {
    if (categoryData.日本語 == keyJP) keyEN = keyJP.replace(keyJP, categoryData.英語)
    return keyEN
  })
}

// 3. オブジェクト内の id を使って、各オブジェクトをグループ化
const exportData = newBookProfiles.reduce((afterData, beforeData) => {
  const targetKey = 'id'
  let addGroupID = beforeData[targetKey] // id をグループ名にする

  afterData[addGroupID] = beforeData // 各オブジェクトにグループ名を追加
  delete beforeData[targetKey] // 各オブジェクトから id を削除
  return afterData
}, {}) // 配列をオブジェクトに変換

// 4．出力
fs.writeFileSync(`json/${fileName}.json`, JSON.stringify(exportData, null, 2));
