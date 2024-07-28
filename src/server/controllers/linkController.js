const HttpError = require('http-errors')
const linkService = require('../services/linkService')
const links = require('../data/links.model').links

module.exports.linkThroughJson = async function (request, _reply) {
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
  const replyData = await linkService.doLinkServiceJson(relayData)

  if (!replyData) {
    throw new HttpError[500]('Command execution failed')
  }

  return {
    replyData
  }
}

module.exports.linkThroughMultipart = async function (request, _reply) {
  const parts = request.parts()
  let serviceId, clientId, segment, token, file, linkData, originalname

  try {
    for await (const part of parts) {
      if (part.fieldname === 'serviceId') {
        serviceId = part.value
        linkData = links.find(link => link.serviceId === serviceId)
        if (!linkData) {
          throw new HttpError[404]('Link not found')
        }
      } else if (part.fieldname === 'clientId') {
        clientId = part.value
      } else if (part.fieldname === 'segment') {
        segment = part.value
      } else if (part.fieldname === 'token') {
        token = part.value
      } else if (part.fieldname === 'file') {
        file = part.file
        originalname = part.filename || 'file'
        const chunks = []
        for await (const chunk of file) {
          chunks.push(chunk)
        }
        file = Buffer.concat(chunks)
      }
    }
  } catch (err) {
    throw new HttpError[400]('Error processing multipart data')
  }

  if (!serviceId || !clientId || !segment || !token || !file) {
    throw new HttpError[400]('Missing required fields')
  }

  const relayData = {
    request,
    linkData,
    clientId,
    segment,
    token,
    file: {
      buffer: file,
      originalname: originalname
    }
  }
  const replyData = await linkService.doLinkServiceMultipart(relayData)

  if (!replyData) {
    throw new HttpError[500]('Command execution failed')
  }

  return {
    replyData
  }
}