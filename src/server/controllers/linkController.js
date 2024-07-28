const HttpError = require('http-errors')
const linkService = require('../services/linkService')
const links = require('../data/links.model').links

module.exports.linkThrough = async function (request, _reply) {
  const { serviceId, clientId, email, token } = request.body
  const linkData = links.find(link => link.serviceId === serviceId)

  if (!linkData) {
    throw new HttpError[404]('Link not found')
  }
  const relayData = {
    request,
    linkData,
    clientId,
    email,
    token
  }
  const replyData = await linkService.doLinkService(relayData)

  if (!replyData) {
    throw new HttpError[500]('Command execution failed')
  }

  return {
    replyData
  }
}
