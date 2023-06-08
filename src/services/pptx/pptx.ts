import type { App } from '@/declarations'
import PptxClass from './pptx.class'

export const registerPptx = (app: App) => {
  app.addService('pptx', new PptxClass(app), {
    hooks: {
      before: {}
    }
  })
}

declare module '../../declarations' {
  // eslint-disable-next-line no-unused-vars
  interface MyServices {
    bars: PptxClass
  }
}
