const fs = require('fs')
let XLSX = require('xlsx')

const fileName = "ExcelToJson"; // ファイル名を入れる（この名前が JSON ファイル名になる）
let workbook = XLSX.readFile(`excel/${fileName}.xlsx`)

// 1．データ取得(JSON)
let bookCategory, bookProfileDirector, bookProfileGroup

workbook.SheetNames.map(sheet => { // 各シート名の変数名を指定
  if ("項目名" == sheet) bookCategory = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
  if ("役職" == sheet) bookProfileDirector = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
  if ("グループ" == sheet) bookProfileGroup = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
})
let bookProfiles = bookProfileDirector.concat(bookProfileGroup) //各 Profile データを結合

// 2. 項目名シートに倣って key 名を英語に変換
let profileInEnglish = bookProfiles.map(profileData => {
  let newProfileData = {}
  for (const key in profileData) {
    if (Object.hasOwnProperty.call(profileData, key)) {
      bookCategory.map(categoryData => {
        if (categoryData.日本語 == key) {
          const newKey = key.replace(key, categoryData.英語)
          newProfileData[newKey] = profileData[key]
        }
      })
    }
  }
  return newProfileData
})

// 3．改行 \n を <br> に変換
profileInEnglish.forEach((profileData) => {
  for (const key in profileData) {
    if (Object.hasOwnProperty.call(profileData, key)) {
      const newElement = profileData[key].replace(/[\n]/g, '<br>')
      profileData[key] = newElement
    }
  }
})

// 4. id を追加しグループ化
const exportData = profileInEnglish.reduce((acc, obj) => {
  let property = 'id'
  let key = obj[property]
  acc[key] = obj
  delete obj[property]
  return acc
}, {})

// 5．出力
fs.writeFileSync(`json/${fileName}.json`, JSON.stringify(exportData, null, 2));
