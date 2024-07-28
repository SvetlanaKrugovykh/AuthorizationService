const Fastify = require('fastify')
const https = require('https')
const authPlugin = require('./plugins/app.auth.plugin')
const fs = require('fs')
const path = require('path')

const credentials = {
  key: fs.readFileSync(path.resolve(__dirname, '../../path/to/localhost.key')),
  cert: fs.readFileSync(path.resolve(__dirname, '../../path/to/localhost.pem'))
}

const app = Fastify({
  trustProxy: true,
  https: credentials
})

app.register(authPlugin)
app.register(require('./routes/auth.route'), { prefix: '/api' })
app.register(require('./routes/links.route'), { prefix: '/api/link' })

module.exports = { app }
