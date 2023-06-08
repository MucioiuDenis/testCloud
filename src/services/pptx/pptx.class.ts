import type { App } from '@/declarations'
import type { Service } from 'zapnode'
import PptxGen from '@ghalex/pptxgenjs'

import { ShapeGenerator } from '@/services/shapeGenerator'
import {
  CriticalPathTable,
  HighFinancialsTable,
  PlanningTable,
  TableGenerator
} from '@/services/tableGenerator'

import {
  convertToObjectArray,
  setCoords,
  sumUpRowPositions
} from '@/services/utils'

const SERVER_URL = 'http://localhost:5000/'

class PptxClass implements Service {
  constructor (protected app: App) {}

  async create (obj: any): Promise<any> {
    const { sheetType, data } = obj
    const pptx = new PptxGen()

    const WIDTH_SLIDE_CM = 33.87
    const HEIGHT_SLIDE_CM = 19.05

    const WIDTH_SLIDE_INCH = WIDTH_SLIDE_CM / 2.54 // 13.33
    const HEIGHT_SLIDE_INCH = HEIGHT_SLIDE_CM / 2.54 // 7.5

    pptx.defineLayout({
      name: 'CUSTOM',
      width: WIDTH_SLIDE_INCH,
      height: HEIGHT_SLIDE_INCH
    })
    pptx.layout = 'CUSTOM'

    console.log(data)
    const slide = pptx.addSlide()
    const shapeGenerator = new ShapeGenerator(pptx, slide)

    const planningTable = new PlanningTable(pptx, slide, data)
    const tableGenerator = new TableGenerator(pptx, slide, data)
    const criticalPathTable = new CriticalPathTable(pptx, slide, data)
    const highFinancials = new HighFinancialsTable(pptx, slide, data)

    shapeGenerator.addHeading('Project name (target)')

    if (data.Related) {
      let oneUnitDay = 0
      let usableHeightTable = 0
      // !ADD TABLE
      if (sheetType === 'Slide5') {
        planningTable.addTable()
        oneUnitDay = planningTable.oneUnitDay
        usableHeightTable = PlanningTable.TABLE_END_POZITION_Y
        shapeGenerator.addSubHeading(
          'Asset Plan showing Primary indication and Life Cycle Innovation'
        )
      } else if (sheetType === 'Slide11') {
        criticalPathTable.addTable()
        oneUnitDay = criticalPathTable.oneUnitDay
        shapeGenerator.addSubHeading('Critical path and near critical path')
      } else if (sheetType === 'Slide8') {
        highFinancials.addTable()
        oneUnitDay = highFinancials.oneUnitDay
        usableHeightTable = HighFinancialsTable.TABLE_END_POZITION_Y - 1.5
        shapeGenerator.addSubHeading(
          'High level project plan to launch with financials'
        )
      } else {
        tableGenerator.addTableFinancialsHight()
        shapeGenerator.addSubHeading('High level plan showing scenarios')
      }

      // ! GENERATE DATE
      const dataGraphic = convertToObjectArray(data, sheetType, oneUnitDay)

      let y = 0
      if (sheetType === 'Slide5') {
        y = sumUpRowPositions(planningTable.rowH, 0)
      } else if (sheetType === 'Slide8') {
        y = sumUpRowPositions(highFinancials.rowH, 0)
      } else {
        y = 1
      }

      let firstElLvlTwo = true
      let numberOfTwoTillNow = 0
      if (sheetType === 'Slide5' || sheetType === 'Slide8') {
        for (let i = 0; i < dataGraphic.length; i++) {
          if (dataGraphic[i].Level === 2 && firstElLvlTwo) {
            firstElLvlTwo = false
            dataGraphic[i].y = y
            numberOfTwoTillNow += 1
          } else if (dataGraphic[i].Level === 2 && !firstElLvlTwo) {
            if (sheetType === 'Slide8') {
              y =
                sumUpRowPositions(highFinancials.rowH, numberOfTwoTillNow) +
                0.3
            } else if (sheetType === 'Slide5') {
              y =
                sumUpRowPositions(planningTable.rowH, numberOfTwoTillNow) + 0.3
            } else {
              y = sumUpRowPositions(planningTable.rowH, numberOfTwoTillNow)
            }
            dataGraphic[i].y = y
            numberOfTwoTillNow += 1
          } else {
            dataGraphic[i].y = y
          }
        }
      }

      const adjustedData = setCoords(dataGraphic, sheetType)

      // !GENERATE SHAPES
      for (let i = 0; i < adjustedData.length; i++) {
        const el = adjustedData[i]

        if (sheetType === 'Slide5' || sheetType === 'Slide8') {
          if (el.Level === 3) {
            shapeGenerator.addRectangle(el, el.x, el.w, el.y)
          }
          if (el.Level === 4) {
            shapeGenerator.addDiamond(el, el.x, el.y - 0.3)
          }
          if (el.Level === 5) {
            shapeGenerator.addDiamondDeadline(
              el,
              el.x,
              el.y - 0.3,
              usableHeightTable
            )
          }
        }

        if (sheetType === 'Slide11') {
          if (el.Level === 3) {
            shapeGenerator.addRectangle(el, el.x, el.w, el.y)
          }
          if (el.Level === 4) {
            shapeGenerator.addDiamond(el, el.x, el.y)
          }
          if (el.Level === 5) {
            shapeGenerator.addDiamondDeadline(
              el,
              el.x,
              el.y,
              usableHeightTable
            )
          }

          const targetElement = adjustedData.find(
            (target: any) => target.Related === el.Predecessor
          )
          if (typeof el.Predecessor === 'number' && targetElement) {
            if (targetElement.Level === 3) {
              const width = targetElement.x - el.x + targetElement.w / 2 - 0.17
              const x = el.x + 0.17
              const y = el.y + 0.085
              const w = width > 0 ? width : Math.abs(width)
              const h = targetElement.y - el.y - 0.085
              const color = 'F36633'
              shapeGenerator.addConnector(el, targetElement, x, y, w, h, color)
            }

            if (targetElement.Level === 4) {
              const x = el.x + 0.17
              const y = el.y + 0.085
              const w = targetElement.x - el.x - 0.085
              const h = targetElement.y - el.y - 0.085
              const color = '959595'
              shapeGenerator.addConnector(el, targetElement, x, y, w, h, color)
            }
          }
        }
      }
    }
    try {
      const res = await pptx.writeFile({ fileName: 'files/test.pptx' })
      return { filename: SERVER_URL + res }
    } catch (err) {
      console.log(err)
      return { err }
    }
  }
}

export default PptxClass
