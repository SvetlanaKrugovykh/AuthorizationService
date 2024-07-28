const HttpError = require('http-errors')
const linkService = require('../services/linkService')

module.exports.linkThrough = async function (request, _reply) {
  const { clientId, email } = request.body
  const payload = { clientId, email }
  const token = await linkService.doLinkService(payload)

  if (!token) {
    throw new HttpError[500]('Command execution failed')
  }

  return {
    token
  }
}
