
const path = require('path')
const fs = require('fs')
const doConvertXlsx = require('../services/xlsxService')
const { uploadFileToDrive } = require('../services/googleDiskService')
require('dotenv').config()


function copyToHttpsOutAndGetUrl(filePath, fileName) {
  const outDir = process.env.HTTPS_OUT || path.join(__dirname, '../temp')
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  const destPath = path.join(outDir, fileName)
  fs.copyFileSync(filePath, destPath)
  const baseUrl = process.env.HTTPS_OUT_URL || 'https://localhost/files/'
  return baseUrl + encodeURIComponent(fileName)
}

exports.linkThroughXlsx = async (req, reply) => {
  try {
    const data = await req.file()
    const variant = req?.headers?.variant || '1'
    const forceLink = req?.headers?.['force-link'] === 'true'
    if (!data) {
      reply.code(400).send({ error: 'No file uploaded' })
      return
    }
    let tempDir = process.env.XLSX_IN || path.join(__dirname, '../temp')
    if (process.env.XLSX_IN) {
      tempDir = path.isAbsolute(process.env.XLSX_IN) ? process.env.XLSX_IN : path.resolve(process.env.XLSX_IN)
    }
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true })
    const filePath = path.join(tempDir, data.filename)
    await new Promise((resolve, reject) => {
      const writeStream = fs.createWriteStream(filePath)
      data.file.pipe(writeStream)
      writeStream.on('finish', resolve)
      writeStream.on('error', reject)
    })

    const convertedFilePath = await doConvertXlsx.doConvertXlsx(filePath)
    const fileName = path.basename(convertedFilePath)
    const stats = fs.statSync(convertedFilePath)
    if (stats.size === 0) {
      reply.code(500).send({ error: 'Converted file is empty' })
      return
    }

    if (variant === '2') {
      try {
        const driveUrl = await uploadFileToDrive(convertedFilePath, fileName)
        reply.send({ url: driveUrl })
      } catch (err) {
        console.error('Google Drive upload failed:', err.message)
        reply.code(500).send({ error: err.message })
      }
    } else if (variant === '1' && forceLink) {
      // Return HTTPS_OUT link
      const url = copyToHttpsOutAndGetUrl(convertedFilePath, fileName)
      reply.send({ url })
    } else if (variant === '1') {
      reply.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      reply.header('Content-Disposition', 'attachment; filename="' + fileName + '"')
      reply.send(fs.createReadStream(convertedFilePath))
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