import dayjs from 'dayjs'
import { uniq } from 'underscore'
export default function convertToObjectArray (
  obj: any,
  sheetType: string,
  oneUnitDay: number
): any[] {
  const yearsData = uniq(
    [...obj.Start, ...obj.Finish].map((date) => dayjs(date).year())
  )

  const minYear = Math.min(...yearsData)
  const startDate = dayjs(`${minYear}-01-01`)
  const keys = Object.keys(obj)
  const length = obj[keys[0]].length // Lungimea array-ului va fi determinată de lungimea unei dintre proprietățile

  const result: any[] = []

  for (let i = 0; i < length; i++) {
    const newObj: any = {} as any
    for (let j = 0; j < keys.length; j++) {
      newObj[keys[j]] = obj[keys[j]][i]
    }
    const data1 = dayjs(obj.Start[i])

    const data2 = dayjs(obj.Finish[i])

    const diferenta = data2.diff(data1, 'day') // returnează diferența în zile

    if (sheetType === 'Slide11' || sheetType === 'Slide8') {
      newObj.x = 1.8 + data1.diff(startDate, 'day') * oneUnitDay // Calculăm valoarea x pe baza datelor
    } else {
      newObj.x = 1.8 + 1 + data1.diff(startDate, 'day') * oneUnitDay // Calculăm valoarea x pe baza datelor
    }
    newObj.w = diferenta * oneUnitDay // Adăugăm proprietatea w cu valoarea 0
    newObj.h = 0 // Adăugăm proprietatea h cu valoarea 0
    newObj.y = 0 // Adăugăm proprietatea y cu valoarea 0

    result.push(newObj)
  }

  return result
}
