// Minimal ZIP reader/writer (store + deflate) with no dependencies.
// Compression uses the platform: CompressionStream/DecompressionStream
// ('deflate-raw') in browsers and Node >= 18, or a faster codec injected
// via setZipCodec() (the Node entry injects node:zlib).

let codec = {
  async inflateRaw(bytes) {
    return streamTransform(bytes, new DecompressionStream('deflate-raw'))
  },
  async deflateRaw(bytes) {
    return streamTransform(bytes, new CompressionStream('deflate-raw'))
  }
}

/**
 * Replace the deflate implementation (e.g. with node:zlib for speed).
 * @param {{inflateRaw(u8: Uint8Array): Promise<Uint8Array>|Uint8Array, deflateRaw(u8: Uint8Array): Promise<Uint8Array>|Uint8Array}} c
 */
export function setZipCodec(c) {
  codec = c
}

async function streamTransform(bytes, transform) {
  const stream = new Blob([bytes]).stream().pipeThrough(transform)
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

// ---------- CRC32 ----------
const CRC_TABLE = (() => {
  const t = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[n] = c >>> 0
  }
  return t
})()

export function crc32(bytes) {
  let crc = 0xffffffff
  for (let i = 0; i < bytes.length; i++) {
    crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8)
  }
  return (crc ^ 0xffffffff) >>> 0
}

const utf8Decoder = new TextDecoder('utf-8')
const utf8Encoder = new TextEncoder()

// ---------- Reader ----------

/**
 * Lists a ZIP archive's entries. File contents are decompressed lazily.
 * @param {Uint8Array} data
 * @returns {Map<string, () => Promise<Uint8Array>>}
 */
export function readZip(data) {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength)
  // Locate End Of Central Directory record (signature 0x06054b50).
  let eocd = -1
  const minPos = Math.max(0, data.length - 22 - 0xffff)
  for (let i = data.length - 22; i >= minPos; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i
      break
    }
  }
  if (eocd < 0) throw new Error('Not a valid .xlsx file (ZIP end record not found)')

  const entryCount = view.getUint16(eocd + 10, true)
  let ptr = view.getUint32(eocd + 16, true)
  if (ptr === 0xffffffff) throw new Error('ZIP64 archives are not supported')

  const entries = new Map()
  for (let n = 0; n < entryCount; n++) {
    if (view.getUint32(ptr, true) !== 0x02014b50) {
      throw new Error('Corrupt ZIP central directory')
    }
    const flags = view.getUint16(ptr + 8, true)
    const method = view.getUint16(ptr + 10, true)
    const compressedSize = view.getUint32(ptr + 20, true)
    const nameLen = view.getUint16(ptr + 28, true)
    const extraLen = view.getUint16(ptr + 30, true)
    const commentLen = view.getUint16(ptr + 32, true)
    const localOffset = view.getUint32(ptr + 42, true)
    const nameBytes = data.subarray(ptr + 46, ptr + 46 + nameLen)
    const name = utf8Decoder.decode(nameBytes)
    ptr += 46 + nameLen + extraLen + commentLen

    if (flags & 0x1) throw new Error('Encrypted (password-protected) files are not supported')
    if (name.endsWith('/')) continue

    entries.set(name, async () => {
      if (view.getUint32(localOffset, true) !== 0x04034b50) {
        throw new Error('Corrupt ZIP local header for ' + name)
      }
      const lNameLen = view.getUint16(localOffset + 26, true)
      const lExtraLen = view.getUint16(localOffset + 28, true)
      const start = localOffset + 30 + lNameLen + lExtraLen
      const raw = data.subarray(start, start + compressedSize)
      if (method === 0) return raw
      if (method === 8) return codec.inflateRaw(raw)
      throw new Error(`Unsupported ZIP compression method ${method} for ${name}`)
    })
  }
  return entries
}

// ---------- Writer ----------

function dosDateTime(date) {
  const time =
    (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2)
  const day =
    ((Math.max(1980, date.getFullYear()) - 1980) << 9) |
    ((date.getMonth() + 1) << 5) |
    date.getDate()
  return { time, day }
}

/**
 * Builds a ZIP archive.
 * @param {{name: string, data: string|Uint8Array}[]} files
 * @returns {Promise<Uint8Array>}
 */
export async function writeZip(files) {
  const { time, day } = dosDateTime(new Date())
  const prepared = await Promise.all(
    files.map(async (f) => {
      const raw = typeof f.data === 'string' ? utf8Encoder.encode(f.data) : f.data
      const compressed = raw.length > 64 ? await codec.deflateRaw(raw) : null
      const useDeflate = compressed && compressed.length < raw.length
      return {
        name: utf8Encoder.encode(f.name),
        crc: crc32(raw),
        size: raw.length,
        method: useDeflate ? 8 : 0,
        payload: useDeflate ? compressed : raw
      }
    })
  )

  let localSize = 0
  let centralSize = 0
  for (const p of prepared) {
    localSize += 30 + p.name.length + p.payload.length
    centralSize += 46 + p.name.length
  }
  const out = new Uint8Array(localSize + centralSize + 22)
  const view = new DataView(out.buffer)
  let offset = 0
  const offsets = []

  for (const p of prepared) {
    offsets.push(offset)
    view.setUint32(offset, 0x04034b50, true)
    view.setUint16(offset + 4, 20, true) // version needed
    view.setUint16(offset + 6, 0x0800, true) // UTF-8 names
    view.setUint16(offset + 8, p.method, true)
    view.setUint16(offset + 10, time, true)
    view.setUint16(offset + 12, day, true)
    view.setUint32(offset + 14, p.crc, true)
    view.setUint32(offset + 18, p.payload.length, true)
    view.setUint32(offset + 22, p.size, true)
    view.setUint16(offset + 26, p.name.length, true)
    view.setUint16(offset + 28, 0, true)
    out.set(p.name, offset + 30)
    out.set(p.payload, offset + 30 + p.name.length)
    offset += 30 + p.name.length + p.payload.length
  }

  const centralStart = offset
  prepared.forEach((p, i) => {
    view.setUint32(offset, 0x02014b50, true)
    view.setUint16(offset + 4, 20, true) // version made by
    view.setUint16(offset + 6, 20, true) // version needed
    view.setUint16(offset + 8, 0x0800, true)
    view.setUint16(offset + 10, p.method, true)
    view.setUint16(offset + 12, time, true)
    view.setUint16(offset + 14, day, true)
    view.setUint32(offset + 16, p.crc, true)
    view.setUint32(offset + 20, p.payload.length, true)
    view.setUint32(offset + 24, p.size, true)
    view.setUint16(offset + 28, p.name.length, true)
    // extra len, comment len, disk no, internal attrs = 0
    view.setUint32(offset + 38, 0, true) // external attrs
    view.setUint32(offset + 42, offsets[i], true)
    out.set(p.name, offset + 46)
    offset += 46 + p.name.length
  })

  view.setUint32(offset, 0x06054b50, true)
  view.setUint16(offset + 8, prepared.length, true)
  view.setUint16(offset + 10, prepared.length, true)
  view.setUint32(offset + 12, offset - centralStart, true)
  view.setUint32(offset + 16, centralStart, true)
  return out
}
