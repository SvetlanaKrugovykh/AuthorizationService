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

    const startTime = Date.now()
    const response = await axios.post(linkData.url, filteredBody, {
      headers: {
        'Content-Type': 'application/json'
      }
    })
    const endTime = Date.now()
    console.log(`${endTime}: JSON request duration: ${endTime - startTime}ms`)

    return response.data
  } catch (err) {
    return null
  }
}


module.exports.doLinkServiceMultipart = async function (relayData) {
  try {
    const { linkData, clientId: relayClientId, token: relayToken, file } = relayData
    const segment_number = relayData?.segment || '1'

    const secretKey = getSecretKey()
    const jwtData = jwt.verify(relayToken, secretKey)
    if (jwtData.clientId !== relayClientId || jwtData.serviceId !== linkData.serviceId) {
      return null
    }

    const formData = new FormData()
    if (file) {
      formData.append('file', file.buffer, { filename: file.originalname })
    }
    formData.append('segment_number', segment_number)
    formData.append('clientId', relayClientId)

    const startTime = Date.now()
    const response = await axios.post(linkData.url, formData, {
      headers: {
        'Content-Type': `multipart/form-data; boundary=${formData._boundary}`
      }
    })
    const endTime = Date.now()
    console.log(`${endTime}: Multipart request duration: ${endTime - startTime}ms`)

    return response.data
  } catch (err) {
    return null
  }
}