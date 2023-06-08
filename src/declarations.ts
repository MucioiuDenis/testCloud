import { type Axios } from 'axios'
import type { Application, Service } from 'zapnode'
import type { AppExpress, AppConfig } from 'zapnode-plugins'

type Services = Record<string, Service>

export interface MyServices extends Services {}

export interface ConfigData {
  name: string
  version: string
  port: number
}

export interface Config extends ConfigData {
  util: {
    getEnv: (name: string) => string
  }
}

export type AppWithConfig = AppConfig<ConfigData>

export type AppExtend = Application<MyServices> & AppExpress & AppWithConfig
export interface App extends AppExtend {
}

declare module 'zapnode' {
  // eslint-disable-next-line no-unused-vars
  interface Params {
    user?: any
  }
}
