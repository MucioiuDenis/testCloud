import './alias'
import { zapnode } from 'zapnode'
import { config, hooks, express as expressPlg } from 'zapnode-plugins'
import registerServices from './services'
import express from 'express'
import type { App } from './declarations'

const expressApp = express()
expressApp.use(express.urlencoded({ extended: true }))
expressApp.use(express.json({ limit: '100kb' }))

expressApp.use('/files', express.static('files'))

const app: App = zapnode({
  plugins: [
    config(),
    hooks(),
    expressPlg(expressApp)
  ]
})

const start = async () => {
  registerServices(app)
  await app.listen(app.config.port)
}

start()
  .then(() => {
    console.info('Server ready at http://localhost:%s', app.config.port)
  })
  .catch(err => {
    console.error(err)
  })
