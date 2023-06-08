import dayjs from 'dayjs'
import { uniq } from 'underscore'
export default class CriticalPathTable {
  private readonly pptx: any
  private readonly slide: any

  static readonly WIDTH_SLIDE_CM = 33.87
  static readonly HEIGHT_SLIDE_CM = 19.05

  static readonly WIDTH_SLIDE_INCH = CriticalPathTable.WIDTH_SLIDE_CM / 2.54
  static readonly HEIGHT_SLIDE_INCH = CriticalPathTable.HEIGHT_SLIDE_CM / 2.54

  static readonly TABLE_START_POZITION_X = 1.8
  static readonly TABLE_END_POZITION_X = 11.8

  static readonly TABLE_START_POZITION_Y = 1.3
  static readonly TABLE_END_POZITION_Y = 6.9

  static readonly FIRST_ROW_HEIGHT = 0.3

  static readonly TABLE_USE_GRAPHIC_Y =
    CriticalPathTable.TABLE_END_POZITION_Y -
    CriticalPathTable.TABLE_START_POZITION_Y -
    CriticalPathTable.FIRST_ROW_HEIGHT

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
        border: { type: 'solid', color: '7f7f7f', pt: 0.2 }
      }
    })

    for (let i = 0; i < data.Activity?.length; i++) {
      if (data.Level[i] === 2) {
        results.push([...ArrayFill])
      }
    }

    return results.length ? results : [ArrayFillNoLvlTwo]
  }

  #getOneUnitDay () {
    const dates = [...this.data.Start, ...this.data.Finish].map((el) =>
      dayjs(el).year()
    )
    dates.sort((a, b) => a - b)

    const minDay = dayjs(`${dates[0]}-01-01`)
    const maxDay = dayjs(`${dates[dates.length - 1]}-12-29`)
    const totalDays = maxDay.diff(minDay, 'days')

    return 10 / totalDays
  }

  addTable () {
    this.rows.push(this.#getRows(this.data))
    this.rows.push(...this.#constructDescriptionRows(this.data, this.rows))

    this.oneUnitDay = this.#getOneUnitDay()

    this.slide.addShape(this.pptx.ShapeType.line, {
      align: 'center',
      x: CriticalPathTable.TABLE_START_POZITION_X,
      y: CriticalPathTable.TABLE_END_POZITION_Y,
      w: 10,
      h: 0,
      line: { color: '7f7f7f', width: 1 }
    })

    this.slide.addShape(this.pptx.ShapeType.line, {
      align: 'center',
      x: CriticalPathTable.TABLE_START_POZITION_X,
      y:
        CriticalPathTable.TABLE_START_POZITION_Y +
        CriticalPathTable.FIRST_ROW_HEIGHT,
      w: 0,
      h: CriticalPathTable.TABLE_USE_GRAPHIC_Y,
      line: { color: '7f7f7f', width: 1 }
    })

    this.slide.addShape(this.pptx.ShapeType.line, {
      align: 'center',
      x: CriticalPathTable.TABLE_END_POZITION_X,
      y:
        CriticalPathTable.TABLE_START_POZITION_Y +
        CriticalPathTable.FIRST_ROW_HEIGHT,
      w: 0,
      h: CriticalPathTable.TABLE_USE_GRAPHIC_Y,
      line: { color: '7f7f7f', width: 1 }
    })

    this.slide.addTable(this.rows, {
      rowH: [
        CriticalPathTable.FIRST_ROW_HEIGHT,
        CriticalPathTable.TABLE_USE_GRAPHIC_Y
      ],
      x: CriticalPathTable.TABLE_START_POZITION_X,
      y: CriticalPathTable.TABLE_START_POZITION_Y,
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
