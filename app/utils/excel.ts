import {
  type WorkBook,
  type WorkSheet,
  read,
  utils,
  Sheet2JSONOpts,
  writeFile,
} from 'xlsx'

export const EXCEL_ACCEPT =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel'

export function importExcelFromBuffer(excelRcFileBuffer: ArrayBuffer): {
  [sheet: string]: WorkSheet
} {
  const workbook = read(excelRcFileBuffer, { type: 'buffer' })
  return workbook.Sheets
}

export function exportExcelFile(
  data: any[],
  sheetName: string,
  fileName: string,
) {
  const jsonWorkSheet = utils.aoa_to_sheet(data)
  const workBook: WorkBook = {
    SheetNames: [sheetName],
    Sheets: {
      [sheetName]: jsonWorkSheet,
    },
  }
  return writeFile(workBook, fileName)
}

export function sheet2JSON<T = Record<string, string>>(
  worksheet: WorkSheet,
  header?: Sheet2JSONOpts['header'],
) {
  return utils.sheet_to_json<T>(worksheet, { raw: true, header })
}

export function splitString(str: string) {
  const result = []
  let currentWord = ''
  let isInParentheses = false

  for (let i = 0; i < str.length; i++) {
    const char = str[i]

    if (char === '(') {
      isInParentheses = true
    } else if (char === ')') {
      isInParentheses = false
    }

    if (char === ',' && !isInParentheses && str[i + 1] === ' ') {
      result.push(currentWord.trim())
      currentWord = ''
      i++ // 跳过空格
    } else {
      currentWord += char
    }
  }

  result.push(currentWord.trim())

  return result
}
