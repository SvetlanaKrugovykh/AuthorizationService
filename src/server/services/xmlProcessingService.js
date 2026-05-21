const { XMLParser } = require('fast-xml-parser')

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  textNodeName: '#text',
  trimValues: true,
  parseTagValue: false
})

function asArray(value) {
  if (value === undefined || value === null) return []
  return Array.isArray(value) ? value : [value]
}

function collectText(node) {
  if (node === null || node === undefined) return ''
  if (typeof node === 'string' || typeof node === 'number' || typeof node === 'boolean') return String(node)
  if (Array.isArray(node)) return node.map(collectText).join(' ').replace(/\s+/g, ' ').trim()

  const parts = []
  for (const [key, value] of Object.entries(node)) {
    if (key === '#text') {
      parts.push(String(value))
      continue
    }
    if (typeof value === 'object') {
      const nested = collectText(value)
      if (nested) parts.push(nested)
    }
  }

  return parts.join(' ').replace(/\s+/g, ' ').trim()
}

function walk(node, visitor) {
  if (node === null || node === undefined) return
  if (Array.isArray(node)) {
    node.forEach(item => walk(item, visitor))
    return
  }
  if (typeof node !== 'object') return

  visitor(node)
  for (const value of Object.values(node)) {
    walk(value, visitor)
  }
}

function getEnvelopeAndTemplate(root) {
  return {
    envelope: root.envelope || {},
    template: root.template || {}
  }
}

function getEnvelopeFieldMap(envelope) {
  const docNode = envelope.document || {}
  const fields = asArray(docNode.field)
  const map = {}
  for (const field of fields) {
    if (!field?.name) continue
    map[field.name] = collectText(field)
  }
  return map
}

function getTemplateTitle(template) {
  return collectText(template?.document?.info?.title) || null
}

function getDuplicateDefinitionMap(template) {
  const map = {}
  walk(template, node => {
    const duplicates = asArray(node.duplicate)
    for (const duplicate of duplicates) {
      if (!duplicate?.name) continue
      map[duplicate.name] = {
        relatedTo: duplicate.relatedTo || null,
        documentId: duplicate.documentId || null
      }
    }
  })
  return map
}

function resolveDuplicateFields(fieldMap, duplicateMap) {
  const resolved = {}
  for (const [name, value] of Object.entries(fieldMap)) {
    const duplicateDef = duplicateMap[name]
    if (!duplicateDef?.relatedTo) continue
    resolved[duplicateDef.relatedTo] = value
  }
  return resolved
}

function normalizeClause(value) {
  if (!value) return null
  return String(value).replace(/\s+/g, ' ').trim().replace(/\.+$/, '')
}

function parseNumericValue(rawValue) {
  if (!rawValue) {
    return { raw: null, value: null, unit: null }
  }

  const raw = String(rawValue).replace(/&#x[0-9a-f]+;|&#\d+;/gi, ' ').replace(/\s+/g, ' ').trim()
  const normalized = raw.replace(/\s+/g, '').replace(',', '.')
  const numeric = Number(normalized.replace(/[^\d.-]/g, ''))
  const value = Number.isFinite(numeric) ? numeric : null

  let unit = null
  if (raw.includes('%')) unit = 'percent'
  if (/грн\.?/i.test(raw)) unit = 'uah'

  return { raw, value, unit }
}

function extractRatesFromAppendixTemplate(template, fieldMap) {
  const rates = []

  walk(template, node => {
    const rows = asArray(node.tr)
    for (const row of rows) {
      const tds = asArray(row?.td)
      if (tds.length < 3) continue

      const clauseText = normalizeClause(collectText(tds[1]))
      if (!clauseText || !/^\d+(\.\d+)*$/.test(clauseText)) continue

      let sourceFieldName = null
      walk(tds[2], tdNode => {
        const fields = asArray(tdNode.field)
        for (const field of fields) {
          if (field?.name && !sourceFieldName) sourceFieldName = field.name
        }
      })

      const rawValue = sourceFieldName
        ? fieldMap[sourceFieldName] || null
        : collectText(tds[2]) || null

      const parsed = parseNumericValue(rawValue)
      rates.push({
        clause: clauseText,
        sourceFieldName,
        valueRaw: parsed.raw,
        value: parsed.value,
        unit: parsed.unit
      })
    }
  })

  const unique = []
  const seen = new Set()
  for (const rate of rates) {
    const key = `${rate.clause}::${rate.sourceFieldName || ''}::${rate.valueRaw || ''}`
    if (seen.has(key)) continue
    seen.add(key)
    unique.push(rate)
  }
  return unique
}

function normalizeCounterparty(raw) {
  return {
    name: raw['Назва постачальника'] || null,
    edrpou: raw['код ЄДРПОУ постачальника'] || null,
    ipn: raw['ІПН постачальника'] || null,
    legalAddress: raw['Адреса постачальника'] || null,
    signer: {
      name: raw['ПІБ підписанта'] || raw['ПІБ'] || null,
      position: raw['Посада підписанта постачальника'] || raw['Посада постачальника'] || null,
      basis: raw['Підстава'] || null
    },
    bankDetails: {
      account: raw['Рахунок постачальника'] || null,
      bankName: raw['Банк постачальника'] || null,
      mfo: raw['МФО банка постачальника'] || null
    },
    taxStatus: raw['Платник податку'] || null,
    gln: raw['GLN Постачальника'] || null
  }
}

function parseContractDocument(root) {
  const { envelope, template } = getEnvelopeAndTemplate(root)
  const fieldMap = getEnvelopeFieldMap(envelope)

  return {
    type: 'contract',
    metadata: {
      createdAt: root.created || null,
      completedAt: root.completed || null,
      envelopeId: envelope.id || null,
      envelopeDocumentId: envelope?.document?.id || null,
      templateId: template.id || null,
      templateVersion: template.version || null,
      templateTitle: getTemplateTitle(template)
    },
    normalized: {
      contract: {
        serviceContractNumber: fieldMap['Номер Договору про надання послуг'] || null,
        serviceContractDate: fieldMap['Дата реєстрації Договору'] || null,
        supplyContractNumber: fieldMap['Номер Договору Поставки'] || null,
        supplyContractDate: fieldMap['Дата реєстрації договору поставки'] || null,
        trademark: fieldMap['Торговельна Марка'] || null
      },
      entities: {
        counterparty: normalizeCounterparty(fieldMap)
      }
    },
    raw: {
      fields: fieldMap
    }
  }
}

function parseAppendix2Document(root) {
  const { envelope, template } = getEnvelopeAndTemplate(root)
  const fieldMap = getEnvelopeFieldMap(envelope)
  const duplicateDefinitionMap = getDuplicateDefinitionMap(template)
  const resolvedDuplicateFields = resolveDuplicateFields(fieldMap, duplicateDefinitionMap)
  const mergedFieldMap = { ...fieldMap, ...resolvedDuplicateFields }

  return {
    type: 'appendix2',
    metadata: {
      createdAt: root.created || null,
      completedAt: root.completed || null,
      envelopeId: envelope.id || null,
      envelopeDocumentId: envelope?.document?.id || null,
      templateId: template.id || null,
      templateVersion: template.version || null,
      templateTitle: getTemplateTitle(template)
    },
    normalized: {
      references: {
        serviceContractNumber: mergedFieldMap['Номер Договору про надання послуг'] || null,
        serviceContractDate: mergedFieldMap['Дата реєстрації Договору'] || null
      },
      entities: {
        counterparty: normalizeCounterparty(mergedFieldMap)
      },
      rates: extractRatesFromAppendixTemplate(template, fieldMap)
    },
    raw: {
      fields: fieldMap,
      duplicateDefinitions: duplicateDefinitionMap,
      resolvedDuplicateFields
    }
  }
}

function parseXmlByType(xmlString, fileType) {
  if (!xmlString || !String(xmlString).trim()) {
    throw new Error('XML payload is empty')
  }

  const parsed = parser.parse(xmlString)
  const root = parsed?.document
  if (!root) {
    throw new Error('Unexpected XML format: root <document> node is missing')
  }

  if (fileType === 'contract') return parseContractDocument(root)
  if (fileType === 'appendix2') return parseAppendix2Document(root)

  throw new Error(`Unsupported fileType: ${fileType}`)
}

module.exports.processXmlBuffer = function processXmlBuffer(fileBuffer, fileType, options = {}) {
  const xmlString = Buffer.isBuffer(fileBuffer)
    ? fileBuffer.toString('utf8')
    : String(fileBuffer || '')

  const result = parseXmlByType(xmlString, fileType)

  return {
    fileType,
    sourceFileName: options.sourceFileName || null,
    processedAt: new Date().toISOString(),
    result
  }
}
