export default function sumUpRowPositions (
  tableRowH: number[],
  position: number
): number {
  if (position < 0 || position >= tableRowH.length) {
    throw new Error('Invalid position')
  }

  let sum = 0
  for (let i = 0; i <= position; i++) {
    sum += tableRowH[i]
  }
  if (tableRowH.length === 2) {
    return 1.8 + sum
  } else if (tableRowH[1] > 1.3 && tableRowH[2] > 1.3) {
    return 1.8 + sum
  } else return 1.5 + sum
}
