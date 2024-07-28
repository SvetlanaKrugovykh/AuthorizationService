const HttpError = require('http-errors')
const authService = require('../services/authService')

module.exports.createAccessToken = async function (request, _reply) {
  const { serviceId, clientId, email } = request.body
  const payload = { serviceId, clientId, email }
  const token = await authService.createAccessToken(payload)

  if (!token) {
    throw new HttpError[500]('Command execution failed')
  }

  return {
    token
  }
}

module.exports.checkAccessToken = async function (request, _reply) {
  const { token } = request.body
  const decodedToken = await authService.checkAccessToken(token)

  const DOMAIN = process.env.DOMAIN || 'localhost'
  const addControl = decodedToken?.email?.includes(DOMAIN) || false;
  if (!decodedToken || !addControl) {
    throw new HttpError[401]('Invalid token: Authorization failed')
  } else {
    return {
      decodedToken
    }
  }
}