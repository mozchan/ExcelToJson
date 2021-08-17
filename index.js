const fs = require('fs')
let XLSX = require('xlsx')
let workbook = XLSX.readFile('excel/ExcelToJson.xlsx')

// 1．データ取得(JSON)
let bookProfile, bookCategory
workbook.SheetNames.map(sheet => {
  if("記入" == sheet) bookProfile = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
  if("項目名" == sheet) bookCategory = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
})

// 2. key 名を英語に変換
const profileInEnglish = bookProfile.map(profileData => {
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

// 3．\n を <br> に変換
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
fs.writeFileSync('json/output.json', JSON.stringify(exportData, null, 2));
