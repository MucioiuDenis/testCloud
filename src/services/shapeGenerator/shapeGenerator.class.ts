export default class ShapeGenerator {
  private readonly pptx: any
  private readonly slide: any

  static readonly WIDTH_SLIDE_CM = 33.87
  static readonly HEIGHT_SLIDE_CM = 19.05

  static readonly WIDTH_SLIDE_INCH = ShapeGenerator.WIDTH_SLIDE_CM / 2.54
  static readonly HEIGHT_SLIDE_INCH = ShapeGenerator.HEIGHT_SLIDE_CM / 2.54

  static readonly TABLE_START_POZITION_X = 1.8
  static readonly TABLE_START_POZITION_Y = 1.3
  static readonly TABLE_END_POZITION_Y = 6.9

  constructor (pptxInstance: any, slideInstance: any) {
    this.pptx = pptxInstance
    this.slide = slideInstance
  }

  addRectangle (data: any, x: number, w: number, y: number) {
    this.slide.addText(`${data.Description}`, {
      shape: this.pptx.ShapeType.roundRect,
      sId: data.Related,
      h: 0.3,
      w,
      x,
      y,
      fill: { color: 'ffffff' },
      line: { color: data.Color, width: 2 },
      align: 'left',
      fontSize: 8,
      fontFace: 'Arial',
      bold: true
    })
  }

  addDiamond (data: any, x: number, y: number) {
    this.slide.addText('', {
      shape: this.pptx.ShapeType.diamond,
      sId: data.Related,
      h: 0.17,
      w: 0.17,
      x,
      y,
      fill: { color: data.Color },
      line: { color: data.Color, width: 1 }
    })

    this.slide.addText(`${data.Description}`, {
      x: x + 0.09, // position of diamond + width of diamond + some gap
      y,
      fontSize: 8,
      fontFace: 'Arial (Body)'
    })
  }

  addDiamondDeadline (data: any, x: number, y: number, tableUseY: number) {
    this.slide.addShape(this.pptx.ShapeType.lineInv, {
      align: 'center',
      x: x + 0.089,
      y,
      w: 0,
      h: tableUseY - y,
      line: { color: 'F36633', width: 1.5, dashType: 'dash' }
    })

    this.slide.addText('', {
      shape: this.pptx.ShapeType.diamond,
      h: 0.17,
      w: 0.17,
      x,
      y,
      fill: { color: data.Color },
      line: { color: data.Color, width: 1 }
    })

    this.slide.addText(`${data.Description}`, {
      x: x + 0.11, // position of diamond + width of diamond + some gap
      y: y + 0.05,
      fontSize: 8,
      fontFace: 'Arial (Body)'
    })
  }

  addConnector (
    data: any,
    targetElement: any,
    x: number,
    y: number,
    w: number,
    h: number,
    color: string | number
  ) {
    this.slide.addShape(this.pptx.ShapeType.line, {
      sId: data.Related,
      x,
      y,
      w,
      h,
      line: {
        color,
        width: 1.5,
        sourceId: data.Related,
        targetId: targetElement.Related,
        sourceAnchorPos: 3,
        targetAnchorPos: 0,
        endArrowType: 'triangle'
      }
    }) // connector test start pos 0 / end pos 1
  }

  addSubHeading (text: string) {
    this.slide.addText(text, {
      x: 0.5,
      y: 1,
      fontFace: 'Arial (Body)',
      fontSize: 24,
      color: '000000'
    })
  }

  addHeading (text: string) {
    this.slide.addText(text, {
      x: 0.5,
      y: 0.5,
      fontFace: 'Arial (Header)',
      fontSize: 24,
      color: 'f36633',
      bold: true
    })
  }
}
