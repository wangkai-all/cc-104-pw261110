/**
 * Minimal ZIP reader for DOCX containers.
 * Supports stored (0) and deflate (8) entries. No ZIP64, no encryption.
 */
import { inflateRawSync } from 'node:zlib';

const EOCD_SIG = 0x06054b50;
const CEN_SIG = 0x02014b50;
const LOC_SIG = 0x04034b50;

function findEocd(buf) {
  const stop = Math.max(0, buf.length - 0xffff - 22);
  for (let i = buf.length - 22; i >= stop; i--) {
    if (buf.readUInt32LE(i) === EOCD_SIG) return i;
  }
  return -1;
}

export function readZip(buf) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  if (buf.length < 22) throw new Error('文件不是有效的 DOCX（内容过短）');
  const eocd = findEocd(buf);
  if (eocd < 0) throw new Error('文件不是有效的 DOCX（找不到 ZIP 目录）');

  const count = buf.readUInt16LE(eocd + 10);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  if (cdOffset === 0xffffffff || count === 0xffff) throw new Error('暂不支持 ZIP64 格式的 DOCX');

  const entries = new Map();
  let p = cdOffset;
  for (let i = 0; i < count; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== CEN_SIG) {
      throw new Error('DOCX 的 ZIP 目录损坏');
    }
    const method = buf.readUInt16LE(p + 10);
    const csize = buf.readUInt32LE(p + 20);
    const usize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
    if (csize === 0xffffffff || usize === 0xffffffff || localOffset === 0xffffffff) {
      throw new Error('暂不支持 ZIP64 格式的 DOCX');
    }
    entries.set(name, { method, csize, usize, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }

  return {
    has(name) {
      return entries.has(name);
    },
    list() {
      return [...entries.keys()];
    },
    read(name) {
      const entry = entries.get(name);
      if (!entry) return null;
      const lh = entry.localOffset;
      if (lh + 30 > buf.length || buf.readUInt32LE(lh) !== LOC_SIG) {
        throw new Error(`DOCX 条目 ${name} 的本地头损坏`);
      }
      const nameLen = buf.readUInt16LE(lh + 26);
      const extraLen = buf.readUInt16LE(lh + 28);
      const start = lh + 30 + nameLen + extraLen;
      const raw = buf.subarray(start, start + entry.csize);
      if (entry.method === 0) {
        if (entry.usize > 200 * 1024 * 1024) throw new Error('DOCX 条目过大');
        return Buffer.from(raw);
      }
      if (entry.method === 8) {
        if (entry.usize > 200 * 1024 * 1024) throw new Error('DOCX 条目过大');
        return inflateRawSync(raw, { maxOutputLength: 200 * 1024 * 1024 });
      }
      throw new Error(`DOCX 条目 ${name} 使用了不支持的压缩方式 ${entry.method}`);
    },
  };
}
