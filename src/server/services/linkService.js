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
    const { linkData, clientId: relayClientId, token: relayToken, file } = relayData

    const secretKey = getSecretKey()
    const jwtData = jwt.verify(relayToken, secretKey)
    if (jwtData.clientId !== relayClientId || jwtData.serviceId !== linkData.serviceId) {
      return null
    }

    const formData = new FormData()
    if (file) {
      formData.append('file', file.buffer, { filename: file.originalname })
    }

    const response = await axios.post(linkData.url, formData, {
      headers: {
        'Content-Type': `multipart/form-data; boundary=${formData._boundary}`
      }
    })

    return response.data
  } catch (err) {
    return null
  }
}