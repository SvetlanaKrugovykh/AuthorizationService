
const path = require('path')
const fs = require('fs')
const doConvertXlsx = require('../services/xlsxService')
require('dotenv').config()


exports.linkThroughXlsx = async (req, reply) => {
  try {
    const data = await req.file()
    const variant = req.body?.variant || '1'
    if (!data) {
      reply.code(400).send({ error: 'No file uploaded' })
      return
    }
    // Save file to input folder
    let tempDir = process.env.XLSX_IN || path.join(__dirname, '../temp')
    // If XLSX_IN is set, ensure absolute path is created
    if (process.env.XLSX_IN) {
      tempDir = path.isAbsolute(process.env.XLSX_IN) ? process.env.XLSX_IN : path.resolve(process.env.XLSX_IN)
    }
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true })
    const filePath = path.join(tempDir, data.filename)
    await data.toFile(filePath)

    if (variant === '1') {
      // Option 1: return file
      const convertedFilePath = await doConvertXlsx.doConvertXlsx(filePath)
      reply.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      reply.send(fs.createReadStream(convertedFilePath))
    } else if (variant === '2') {
      // Option 2: return Google Drive link (stub for now)
      const driveUrl = 'https://drive.google.com/link/' + data.filename
      reply.send({ url: driveUrl })
    } else {
      reply.code(400).send({ error: 'Unknown variant' })
    }
  } catch (err) {
    reply.code(500).send({ error: err.message })
  }
}
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
  let serviceId, clientId, segment_number, token, file, linkData, originalname

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
      } else if (part.fieldname === 'segment_number') {
        segment_number = part.value
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

  // Check required fields
  if (!serviceId || !clientId || !segment_number || !token || !file) {
    throw new HttpError[400]('Missing required fields')
  }

  const relayData = {
    request,
    linkData,
    clientId,
    segment_number,
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