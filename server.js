import express from 'express'
import cors from 'cors'
import path from 'path'
import { fileURLToPath } from 'url'
import fs from 'fs'

const app = express()
const port = Number(process.env.PORT) || 4000
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const distPath = path.join(__dirname, 'dist')
const indexHtmlPath = path.join(distPath, 'index.html')

const serviceStates = {
  shipment: false,
  reports: true,
  inventory: true,
}

app.use(cors({ origin: true }))
app.use(express.json())
app.use('/api', (_req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate')
  res.set('Pragma', 'no-cache')
  res.set('Expires', '0')
  next()
})

app.get('/api/services', (_req, res) => {
  res.json({ services: serviceStates })
})

app.get('/api/services/:serviceName', (req, res) => {
  const { serviceName } = req.params
  const enabled = serviceStates[serviceName] ?? true

  res.json({
    service: serviceName,
    enabled,
  })
})

app.post('/api/services/:serviceName', (req, res) => {
  const { serviceName } = req.params
  const { enabled } = req.body ?? {}

  if (typeof enabled !== 'boolean') {
    return res.status(400).json({ message: 'Field "enabled" must be a boolean.' })
  }

  serviceStates[serviceName] = enabled

  return res.json({
    service: serviceName,
    enabled,
  })
})

if (fs.existsSync(distPath)) {
  app.use(express.static(distPath))

  app.get(/^\/(?!api).*/, (req, res) => {
    if (req.path.startsWith('/api')) {
      return res.status(404).json({ message: 'API route not found.' })
    }

    return res.sendFile(indexHtmlPath)
  })
}

app.listen(port, '0.0.0.0', () => {
  console.log(`API server running on http://localhost:${port}`)
})
