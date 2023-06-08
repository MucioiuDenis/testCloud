import dayjs from 'dayjs'
import { uniq } from 'underscore'
export default class HighFinancialsTable {
  private readonly pptx: any
  private readonly slide: any

  static readonly WIDTH_SLIDE_CM = 33.87
  static readonly HEIGHT_SLIDE_CM = 19.05

  static readonly WIDTH_SLIDE_INCH = HighFinancialsTable.WIDTH_SLIDE_CM / 2.54
  static readonly HEIGHT_SLIDE_INCH =
    HighFinancialsTable.HEIGHT_SLIDE_CM / 2.54

  static readonly TABLE_START_POZITION_X = 1.8
  static readonly TABLE_END_POZITION_X = 11.8

  static readonly TABLE_START_POZITION_Y = 1.3
  static readonly TABLE_END_POZITION_Y = 7.1

  static readonly FIRST_ROW_HEIGHT = 0.3

  static readonly TABLE_USE_GRAPHIC_Y =
    HighFinancialsTable.TABLE_END_POZITION_Y -
    HighFinancialsTable.TABLE_START_POZITION_Y -
    HighFinancialsTable.FIRST_ROW_HEIGHT -
    1.5

  static readonly TABLE_WIDTH = 10

  public rows: any[]
  public rowH: number[]
  public oneUnitDay: number

  private readonly data: any
  constructor (pptxInstance: any, slideInstance: any, data: any) {
    this.pptx = pptxInstance
    this.slide = slideInstance
    this.data = data
    this.rows = []
    this.rowH = []
    this.oneUnitDay = 0
  }

  #divideValue (array: number[], num: number): number[] {
    const result: number[] = []

    const totalThrees = array.filter((v) => v === 3).length
    const unit = num / totalThrees // calculate the unit value

    const twosIndexes = array
      .map((v, i) => (v === 2 ? i : -1))
      .filter((i) => i !== -1)
    twosIndexes.push(array.length) // add an extra index at the end for easier calculations

    for (let i = 0; i < twosIndexes.length - 1; i++) {
      // count number of 3s between current 2 and next 2 (or end of array)
      const threesCount = array
        .slice(twosIndexes[i], twosIndexes[i + 1])
        .filter((v) => v === 3).length

      // calculate the total sum for this segment
      const segmentSum = unit * threesCount

      // add this sum to the result array
      result.push(segmentSum)
    }

    return result.length < 2
      ? [
          HighFinancialsTable.FIRST_ROW_HEIGHT,
          HighFinancialsTable.TABLE_USE_GRAPHIC_Y
        ]
      : [HighFinancialsTable.FIRST_ROW_HEIGHT, ...result]
  }

  #getRows (data: any): any[] {
    // sort and get the uniq years
    const yearsData = uniq(
      [...data.Start, ...data.Finish].map((date) => dayjs(date).year())
    )
    const completedYears: number[] = []

    const minYear = Math.min(...yearsData)
    const maxYear = Math.max(...yearsData)
    for (let i = minYear; i <= maxYear; i++) {
      completedYears.push(i)
    }

    return [...completedYears]
  }

  #constructDescriptionRows (data: any, rows: any[]): any[] {
    const results: any[] = []
    const ArrayFill = Array(rows[0]?.length).fill({
      text: '',
      options: {
        align: 'center',
        fontFace: 'Arial',
        fontSize: 8,
        color: '000000',
        vert: 'wordArtVert',
        border: { type: 'solid', color: '7f7f7f', pt: 0.1 }
      }
    })

    const ArrayFillNoLvlTwo = Array(rows[0]?.length).fill({
      text: '',
      options: {
        align: 'center',
        fontFace: 'Arial',
        fontSize: 8,
        color: '000000',
        vert: 'wordArtVert',
        border: { type: 'solid', color: '7f7f7f', pt: 0.1 }
      }
    })

    for (let i = 0; i < data.Activity?.length; i++) {
      if (data.Level[i] === 2) {
        results.push([...ArrayFill])
      }
    }

    return results.length ? results : [ArrayFillNoLvlTwo]
  }

  #getWidthColumns () {
    // we start from 0 because the first element it's empty and he have a width defined already who is staic
    const columns: any[] = []
    for (let i = 0; i < this.rows[0].length; i++) {
      columns.push(10 / (this.rows[0].length - 1))
    }
    return columns
  }

  #getOneUnitDay () {
    const dates = [...this.data.Start, ...this.data.Finish].map((el) =>
      dayjs(el).year()
    )
    dates.sort((a, b) => a - b)

    const minDay = dayjs(`${dates[0]}-01-01`)
    const maxDay = dayjs('2034-12-29')
    const totalDays = maxDay.diff(minDay, 'days')

    return 10 / totalDays
  }

  addTable () {
    this.rows.push([
      2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034
    ])
    // this.rows.push(this.#getRows(this.data))
    this.rows.push(...this.#constructDescriptionRows(this.data, this.rows))

    const rowH = this.#divideValue(
      this.data.Level,
      HighFinancialsTable.TABLE_USE_GRAPHIC_Y
    )

    const colW = [1, ...this.#getWidthColumns()]
    this.oneUnitDay = this.#getOneUnitDay()
    this.rowH = rowH

    this.rows.push(
      [
        {
          text: 'Cumulative EPE/IPE to Key inflection point',
          options: {
            colW: 3,
            colspan: 3,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 9.5,
            italic: true,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 'XX.X/XX.X',
          options: {
            colW: 2,
            colspan: 2,
            align: 'right',
            fontFace: 'Calibri',
            fontSize: 10.5,

            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 'XX.X/XX.X',
          options: {
            colW: 2,
            colspan: 2,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 10.5,

            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 'XX.X/XX.X',
          options: {
            colW: 2,
            colspan: 2,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 10.5,

            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 'XX.X/XX.X',
          options: {
            colW: 2,
            colspan: 2,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 10.5,

            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: ' ',
          options: {
            colW: 2,
            colspan: 2,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 10.5,

            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        }
      ],
      [
        {
          text: ' Gross EPE (£m) by year',
          options: {
            colW: 3,
            colspan: 3,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 9.5,
            italic: true,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        ...[12.2, 20.4, 56.6, 121.1, 104.5, 86.9, 30.1, 92, 'XX.X'].map(
          (el) => {
            return {
              text: el,
              options: {
                align: 'center',
                fontFace: 'Calibri',
                fontSize: 10.5,
                color: '7f7f7f',
                border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
              }
            }
          }
        )
      ],
      [
        {
          text: ' IPE (£m)  by year',
          options: {
            colW: 3,
            colspan: 3,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 9.5,
            italic: true,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        ...[3.1, 10.2, 4.5, 6.9, 4.3, 6.0, 10.2, 10, 'XX.X'].map((el) => {
          return {
            text: el,
            options: {
              align: 'center',
              fontFace: 'Calibri',
              fontSize: 10.5,
              color: '7f7f7f',
              border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
            }
          }
        })
      ],
      [
        {
          text: ' Milestone payments (£m)',
          options: {
            colW: 3,
            colspan: 3,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 9.5,
            italic: true,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: ' ',
          options: {
            colW: 3,
            colspan: 3,
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 10,
          options: {
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: ' ',
          options: {
            colW: 2,
            colspan: 2,
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 10,
          options: {
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 10,
          options: {
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 'XX.X',
          options: {
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        }
      ],
      [
        {
          text: ' PTRS %',
          options: {
            colW: 3,
            colspan: 3,
            align: 'left',
            fontFace: 'Calibri',
            fontSize: 9.5,
            italic: true,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 20,
          options: {
            colW: 2,
            colspan: 2,
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 50,
          options: {
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 70,
          options: {
            colW: 4,
            colspan: 4,
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 90,
          options: {
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        },
        {
          text: 'XX.X',
          options: {
            align: 'center',
            fontFace: 'Calibri',
            fontSize: 10.5,
            color: '7f7f7f',
            border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
          }
        }
      ]
    )

    this.slide.addShape(this.pptx.ShapeType.line, {
      align: 'center',
      x: HighFinancialsTable.TABLE_START_POZITION_X,
      y: HighFinancialsTable.TABLE_END_POZITION_Y - 0.12,
      w: 10,
      h: 0,
      line: { color: '7f7f7f', width: 1 }
    })

    this.slide.addShape(this.pptx.ShapeType.line, {
      align: 'center',
      x: HighFinancialsTable.TABLE_START_POZITION_X,
      y:
        HighFinancialsTable.TABLE_START_POZITION_Y +
        HighFinancialsTable.FIRST_ROW_HEIGHT,
      w: 0,
      h: HighFinancialsTable.TABLE_USE_GRAPHIC_Y + 1.39,
      line: { color: '7f7f7f', width: 1 }
    })

    this.slide.addShape(this.pptx.ShapeType.line, {
      align: 'center',
      x: HighFinancialsTable.TABLE_END_POZITION_X,
      y:
        HighFinancialsTable.TABLE_START_POZITION_Y +
        HighFinancialsTable.FIRST_ROW_HEIGHT,
      w: 0,
      h: HighFinancialsTable.TABLE_USE_GRAPHIC_Y + 1.39,
      line: { color: '7f7f7f', width: 1 }
    })

    function sumUpToPosition (array: number[], position: number): number {
      if (position < 0 || position >= array.length) {
        throw new Error('Invalid position')
      }

      let sum = 0
      for (let i = 0; i <= position; i++) {
        sum += array[i]
      }

      return 1.3 + sum
    }

    for (let i = 1; i < this.rowH.length; i++) {
      const y = sumUpToPosition(this.rowH, i)
      this.slide.addShape(this.pptx.ShapeType.line, {
        align: 'center',
        x: HighFinancialsTable.TABLE_START_POZITION_X,
        y,
        w: HighFinancialsTable.TABLE_WIDTH,
        h: 0,
        line: { color: '7f7f7f', width: 1 }
      })
    }

    this.slide.addTable(this.rows, {
      colW,
      rowH,
      x: HighFinancialsTable.TABLE_START_POZITION_X,
      y: HighFinancialsTable.TABLE_START_POZITION_Y,
      w: 10,
      align: 'center',
      fontFace: 'Arial',
      fontSize: 11,
      color: '7f7f7f',
      border: { type: 'solid', color: '7f7f7f', pt: 0.9 },
      valign: 'middle'
    })
  }
}
