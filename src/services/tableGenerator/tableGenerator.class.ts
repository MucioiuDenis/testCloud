export default class TableGenerator {
  private readonly rowsFinancials: any = [
    [2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034],
    [
      ...Array(11)
        .fill('')
        .map((el) => {
          return {
            text: ' ',
            options: {
              align: 'right',
              fontFace: 'Calibri',
              valign: 'bottom',
              fontSize: 10,
              italic: true,
              color: '7f7f7f',
              border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
            }
          }
        }),
      {
        text: 'Total (£m)',
        options: {
          // colW: 12,
          // colspan: 12,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'bottom',
          fontSize: 10,
          italic: true,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
        }
      }
    ],
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
      ...[12.2, 20.4, 56.6, 121.1, 104.5, 86.9, 30.1, 92, 'XX.X'].map((el) => {
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
  ]

  private readonly rowsFinancialsHight: any = [
    [
      {
        text: ' Scenario',
        options: {
          align: 'center',
          fontFace: 'Arial',
          valign: 'center',
          fontSize: 11,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      2024,
      2025,
      2026,
      2027,
      2028,
      2029,
      2030,
      2031,
      2032,
      2033,
      ''
    ],
    [
      {
        text: 'Recommended option Launch 20XX',
        options: {
          align: 'center',
          fontFace: 'Georgia',
          valign: 'center',
          italic: true,

          fontSize: 9,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'Total(£m)',
        options: {
          colW: 11,
          colspan: 11,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'bottom',
          fontSize: 11,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: ' Cumulative EPE /IPE (£m)          ',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          colW: 3,
          colspan: 3,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          colW: 5,
          colspan: 5,
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: ' PTRS %',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          colW: 3,
          colspan: 3,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          colW: 5,
          colspan: 5,
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: '',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: 'Recommended option Launch 20XX',
        options: {
          align: 'center',
          fontFace: 'Georgia',
          valign: 'center',
          italic: true,

          fontSize: 9,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'Total(£m)',
        options: {
          colW: 11,
          colspan: 11,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'bottom',
          fontSize: 11,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: ' Cumulative EPE /IPE (£m)          ',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          colW: 3,
          colspan: 3,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          colW: 5,
          colspan: 5,
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: ' PTRS %',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          colW: 3,
          colspan: 3,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          colW: 5,
          colspan: 5,
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: '',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          bold: true,
          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: 'Recommended option Launch 20XX',
        options: {
          align: 'center',
          fontFace: 'Georgia',
          valign: 'center',
          italic: true,

          fontSize: 9,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'Total(£m)',
        options: {
          colW: 11,
          colspan: 11,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'bottom',
          fontSize: 11,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: ' Cumulative EPE /IPE (£m)          ',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          colW: 3,
          colspan: 3,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X/XX.X',
        options: {
          colW: 5,
          colspan: 5,
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX.X',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          bold: true,
          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ],
    [
      {
        text: ' PTRS %',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          colW: 3,
          colspan: 3,
          align: 'right',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          color: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: 'XX',
        options: {
          colW: 5,
          colspan: 5,
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,

          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      },
      {
        text: '',
        options: {
          align: 'center',
          fontFace: 'Calibri',
          valign: 'center',
          fontSize: 10.5,
          bold: true,
          cqolor: '7f7f7f',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      }
    ]
  ]

  private readonly pptx: any
  private readonly slide: any

  static readonly WIDTH_SLIDE_CM = 33.87
  static readonly HEIGHT_SLIDE_CM = 19.05

  static readonly WIDTH_SLIDE_INCH = TableGenerator.WIDTH_SLIDE_CM / 2.54
  static readonly HEIGHT_SLIDE_INCH = TableGenerator.HEIGHT_SLIDE_CM / 2.54

  static readonly TABLE_START_POZITION_X = 1.8
  static readonly TABLE_START_POZITION_Y = 1.3
  static readonly TABLE_END_POZITION_Y = 6.9

  public rows: any[]
  private readonly data: any

  constructor (pptxInstance: any, slideInstance: any, data: any) {
    this.pptx = pptxInstance
    this.slide = slideInstance
    this.data = data
    this.rows = [
      [
        '',
        2023,
        2024,
        2025,
        2026,
        2027,
        2028,
        2029,
        2030,
        2031,
        2032,
        2033,
        2034
      ]
    ]
  }

  addTableFinancials () {
    this.slide.addTable(this.rowsFinancials, {
      rowH: [0.3, 3.9, 0.3, 0.3, 0.3, 0.3],
      x: TableGenerator.TABLE_START_POZITION_X,
      y: TableGenerator.TABLE_START_POZITION_Y,
      w: 10,
      align: 'center',
      fontFace: 'Calibri',
      fontSize: 10,
      color: '7f7f7f',
      border: { type: 'solid', color: '7f7f7f', pt: 1 },
      valign: 'middle'
    })
  }

  addTableFinancialsHight () {
    const colW = Array(11).fill(9 / (this.rowsFinancialsHight[0].length - 1.7))

    this.slide.addTable(this.rowsFinancialsHight, {
      rowH: [0.3, 1.2, 0.3, 0.3, 1.2, 0.3, 0.3, 1.2, 0.3, 0.3],
      colW: [1.8, ...colW],
      x: 0.7,
      y: TableGenerator.TABLE_START_POZITION_Y,
      w: 11,
      align: 'center',
      fontFace: 'Calibri',
      fontSize: 10,
      color: '7f7f7f',
      border: { type: 'solid', color: '7f7f7f', pt: 1 },
      valign: 'middle'
    })
  }

  addTablePlanning () {
    for (let i = 0; i < this.data.Activity?.length; i++) {
      const ArrayFill = Array(this.rows[0]?.length - 1).fill({
        text: '',
        options: {
          align: 'center',
          fontFace: 'Arial',
          fontSize: 8,
          color: '000000',
          vert: 'wordArtVert',
          border: { type: 'solid', color: '7f7f7f', pt: 1 }
        }
      })
      if (this.data.Level[i] === 2) {
        this.rows.push([
          {
            text: this.data.Description[i],
            options: {
              align: 'center',
              fontFace: 'Arial',
              fontSize: 8,
              color: '000000',
              vert: 'wordArtVert',
              border: { type: 'solid', color: '7f7f7f', pt: 0.5 }
            }
          },
          ...ArrayFill
        ])
      }
    }

    const rowH = Array(this.rows?.length - 1).fill(
      5.3 / (this.rows?.length - 1)
    )

    this.slide.addTable(this.rows, {
      rowH: rowH?.length === 0 ? [0.3, 5.3] : [0.3, ...rowH],
      x: TableGenerator.TABLE_START_POZITION_X,
      y: TableGenerator.TABLE_START_POZITION_Y,
      w: 10,
      align: 'center',
      fontFace: 'Arial',
      fontSize: 11,
      color: '7f7f7f',
      border: { type: 'solid', color: '7f7f7f', pt: 1 },
      valign: 'middle'
    })
  }

  calculateOneDayValue () {
    const oZi = (8.95 - 0.7) / ((this.rows[0]?.length - 2) * 365)
    return oZi
  }
}
