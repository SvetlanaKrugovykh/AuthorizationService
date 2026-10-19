const Fastify = require('fastify')
const https = require('https')
const authPlugin = require('./plugins/app.auth.plugin')
const fs = require('fs')
const path = require('path')
require('dotenv').config()

const sslCertDir = process.env.SSL_CERT_DIR || path.resolve(__dirname, '../../path/to')
const credentials = {
  key: fs.readFileSync(path.resolve(sslCertDir, process.env.SSL_KEY_FILE || 'localhost.key')),
  cert: fs.readFileSync(path.resolve(sslCertDir, process.env.SSL_CERT_FILE || 'localhost.pem'))
}

const app = Fastify({
  trustProxy: true,
  logger: true,
  https: credentials
})

app.register(require('@fastify/multipart'))
app.addContentTypeParser('application/json', { parseAs: 'string' }, app.getDefaultJsonParser('ignore', 'ignore'))

app.register(authPlugin)

// Register routes
app.register(require('./routes/auth.route'), { prefix: '/api' })
app.register(require('./routes/links.route'), { prefix: '/api/link' })
app.register(require('./routes/xml.processing.route'), { prefix: '/api/xml' })

module.exports = { app }
