const HttpError = require('http-errors')
const jwt = require('jsonwebtoken')
const { getSecretKey } = require('../guards/getCredentials')

module.exports = function (request, _reply, done) {

  const allowedIps = (process.env.VIP_API_ALLOWED_IPS || '').split(',').map(ip => ip.trim())
  let clientIp = request.ip
  // Normalize IPv6-mapped IPv4 addresses
  if (clientIp.startsWith('::ffff:')) {
    clientIp = clientIp.replace('::ffff:', '')
  }
  console.log(`[is-authorized.guard] Client IP: ${clientIp}, Allowed IPs: ${allowedIps}`)
  if (allowedIps.includes(clientIp)) {
    return done()
  }

  const secretKey = getSecretKey()
  const data = jwt.verify(request.auth.token, secretKey)
  if (!data.clientId) {
    throw new HttpError.Unauthorized('Authorization required')
  }

  done()
}
