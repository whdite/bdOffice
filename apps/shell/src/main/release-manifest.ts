import { createPublicKey, verify } from 'node:crypto'
import { blake2b } from '@noble/hashes/blake2.js'
import currentKey from '../../../../release-keys/bigdogoffice.minisign.pub?raw'
import legacyKey from '../../../../release-keys/bigdogoffice-legacy-through-0.9.0.minisign.pub?raw'

/** Minisign ED packets sign BLAKE2b-512(file). The second signature binds the trusted comment. */
export function verifyReleaseManifest(bytes: Buffer, sidecar: Buffer, legacy: boolean): string {
  const publicKey = (legacy ? legacyKey : currentKey).trim().split(/\r?\n/)[1]!
  const key = Buffer.from(publicKey, 'base64')
  const lines = sidecar.toString('utf8').trimEnd().split(/\r?\n/)
  const packet = Buffer.from(lines[1] ?? '', 'base64')
  if (
    lines.length !== 4 ||
    !lines[2]?.startsWith('trusted comment: ') ||
    packet.length !== 74 ||
    packet.toString('ascii', 0, 2) !== 'ED' ||
    !packet.subarray(2, 10).equals(key.subarray(2, 10))
  )
    throw new Error('Unexpected release signature')
  const signer = createPublicKey({
    key: Buffer.concat([Buffer.from('302a300506032b6570032100', 'hex'), key.subarray(10)]),
    format: 'der',
    type: 'spki',
  })
  const signature = packet.subarray(10)
  const comment = Buffer.from(lines[2].slice('trusted comment: '.length))
  if (
    !verify(null, blake2b(bytes, { dkLen: 64 }), signer, signature) ||
    !verify(null, Buffer.concat([signature, comment]), signer, Buffer.from(lines[3]!, 'base64'))
  )
    throw new Error('Invalid release manifest signature')
  return publicKey
}
