const HttpError = require('http-errors')
const authService = require('../services/authService')

module.exports.createAccessToken = async function (request, _reply) {
  const { clientId, email } = request.body
  const payload = { clientId, email }
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

  if (!decodedToken) {
    throw new HttpError[401]('Invalid token: Authorization failed')
  } else {
    return {
      valid: true, data: decodedToken
    }
  }
}