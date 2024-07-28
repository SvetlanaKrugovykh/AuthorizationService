const jwt = require('jsonwebtoken')
const crypto = require('crypto')
const fs = require('fs')
const { getSecretKey } = require('../guards/getCredentials')

module.exports.doLinkService = async function (token) {
  try {
    const secretKey = getSecretKey()
    return jwt.verify(token, secretKey)
  } catch (err) {
    return null
  }
}

