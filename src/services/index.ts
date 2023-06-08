import type { App } from '../declarations'

import { registerPptx } from './pptx/pptx'

const registerServices = (app: App) => {
  registerPptx(app)
}

export default registerServices
