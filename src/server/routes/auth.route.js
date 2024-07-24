const authController = require('../controllers/authController')
const isAuthorizedGuard = require('../guards/is-authorized.guard')
const authentificationSchema = require('../schemas/authentification.schema')

module.exports = (fastify, _opts, done) => {
  fastify.route({
    method: 'POST',
    url: '/auth/generate-access-token/',
    handler: authController.createAccessToken,
    preHandler: [
      isAuthorizedGuard
    ],
    schema: authentificationSchema
  })

  fastify.route({
    method: 'POST',
    url: '/auth/check-access-token/',
    handler: authController.checkAccessToken,
    preHandler: [
      isAuthorizedGuard
    ],
    schema: authentificationSchema
  })

  done()
}