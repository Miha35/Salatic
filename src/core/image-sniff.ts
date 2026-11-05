/** Magic-number checks for image formats that need our own decoders (kept free of heavy imports). */

export const isTiff = (b: Uint8Array): boolean => b.length > 8 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 0x2a));

/** ISO-BMFF brands of HEIF stills (HEIC = HEVC-coded HEIF). */
export const isHeic = (b: Uint8Array): boolean => {
  if (b.length < 16 || b[4] !== 0x66 || b[5] !== 0x74 || b[6] !== 0x79 || b[7] !== 0x70) return false;
  const size = Math.min(b.length, (b[0] << 24) | (b[1] << 16) | (b[2] << 8) | b[3]);
  const brands: string[] = [];
  for (let p = 8; p + 4 <= size; p += 4) if (p !== 12) brands.push(String.fromCharCode(b[p], b[p + 1], b[p + 2], b[p + 3]));
  // AVIF shares the HEIF container; browsers decode it natively
  if (brands.some((x) => x === 'avif' || x === 'avis')) return false;
  return ['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'mif1', 'msf1'].includes(brands[0]);
};
