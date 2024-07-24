require('dotenv').config()
const { app } = require('./index')
const HOST = process.env.HOST || '0.0.0.0'

app.listen({ port: process.env.PORT || 9876, host: HOST }, (err, address) => {
  if (err) {
    app.log.error(err)
    process.exit(1)
  }

  console.log(`${new Date()}:[API] Service listening on ${address}`)
})

