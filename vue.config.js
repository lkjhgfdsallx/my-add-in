const fs = require('node:fs')
const path = require('node:path')
const homedir = require('node:os').homedir()
const { defineConfig } = require('@vue/cli-service')

module.exports = defineConfig({
  transpileDependencies: true,
  devServer: {
    port: 3000,
    https: {
      key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
      cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
      ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`)),
    },
  },
})
