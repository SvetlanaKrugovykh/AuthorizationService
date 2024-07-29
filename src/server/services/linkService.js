const axios = require('axios')
const jwt = require('jsonwebtoken')
const FormData = require('form-data')
const { getSecretKey } = require('../guards/getCredentials')

module.exports.doLinkServiceJson = async function (relayData) {
  try {
    const { request, linkData, clientId: relayClientId, token: relayToken } = relayData
    const secretKey = getSecretKey()
    const jwtData = jwt.verify(relayToken, secretKey)
    if (jwtData.clientId !== relayClientId || jwtData.serviceId !== linkData.serviceId) {
      return null
    }

    const { serviceId, clientId, email, token, ...filteredBody } = request.body

    const response = await axios.post(linkData.url, filteredBody, {
      headers: {
        'Content-Type': 'application/json'
      }
    })

    return response.data
  } catch (err) {
    return null
  }
}


module.exports.doLinkServiceMultipart = async function (relayData) {
  try {
    const { request, linkData, clientId, email, token, file } = relayData
    const secretKey = getSecretKey()
    const jwtData = jwt.verify(token, secretKey)

    if (jwtData.clientId !== clientId || jwtData.serviceId !== linkData.serviceId) {
      return null
    }

    const formData = new FormData()
    if (file) {
      formData.append('file', file.buffer, { filename: file.originalname })
    }

    const agent = new https.Agent({
      rejectUnauthorized: false
    })

    console.log('linkData.url', linkData.url)
    console.log('request.body', request.body)

    const response = await axios.post(linkData.url, formData, {
      headers: {
        ...request.headers,
        'Content-Type': `multipart/form-data; boundary=${formData._boundary}`
      },
      httpsAgent: agent
    })

    return response.data
  } catch (err) {
    return null
  }
}