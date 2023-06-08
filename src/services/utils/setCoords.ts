import { CriticalPathTable } from '@/services/tableGenerator'

export default function setCoords (arr: any[], sheetType: string): any[] {
  if (sheetType === 'Slide5' || sheetType === 'Slide8') {
    for (let i = 0; i < arr.length - 1; i++) {
      if (arr[i].Level !== 3) {
        continue
      }

      for (let j = i + 1; j < arr.length; j++) {
        if (arr[j].Level !== 3) {
          continue
        }

        if (arr[i].x < arr[j].x + arr[j].w && arr[i].x + arr[i].w > arr[j].x) {
          if (arr[i].y === arr[j].y) {
            if (sheetType === 'Slide8') {
              arr[j].y += 0.4
            } else {
              arr[j].y += 0.5
            }
          }
        }
      }
    }
  }

  if (sheetType === 'Slide11') {
    let y = CriticalPathTable.TABLE_START_POZITION_Y + 0.3
    for (let i = 0; i < arr.length; i++) {
      if (i >= 1 && arr[i - 1].Level === 3 && arr[i].Level === 4) {
        y += 0.4
        arr[i].y = y
      }
      if (i >= 1 && arr[i - 1].Level === 4 && arr[i].Level === 4) {
        y += 0.3
        arr[i].y = y
      }
      if (i >= 1 && arr[i - 1].Level === 3 && arr[i].Level === 3) {
        y += 0.3
        arr[i].y = y
      }
      if (arr[i].Level === 3) {
        y += 0.22
        arr[i].y = y
      }
      if (i === 0 && arr[i].Level === 4) {
        arr[i].y = y + 0.1
      }
    }
  }

  return arr
}
