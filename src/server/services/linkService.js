const axios = require('axios')
const jwt = require('jsonwebtoken')
const crypto = require('crypto')
const fs = require('fs')
const { getSecretKey } = require('../guards/getCredentials')

module.exports.doLinkServiceJson = async function (relayData) {
  try {
    const { request, linkData, clientId, email, token } = relayData
    const secretKey = getSecretKey()
    const jwtData = jwt.verify(token, secretKey)
    if (jwtData.clientId !== clientId || jwtData.serviceId !== linkData.serviceId) {
      return null
    }
    const response = await axios.post(linkData.url, request.body, {
      headers: request.headers
    })

    return response.data;
  } catch (err) {
    return null
  }
}


module.exports.doLinkServiceMultipart = async function (relayData) {
  try {
    const { request, linkData, clientId, email, token } = relayData
    const secretKey = getSecretKey()
    const jwtData = jwt.verify(token, secretKey)
    if (jwtData.clientId !== clientId || jwtData.serviceId !== linkData.serviceId) {
      return null
    }
    const response = await axios.post(linkData.url, request.body, {
      headers: request.headers
    })

    return response.data;
  } catch (err) {
    return null
  }
}