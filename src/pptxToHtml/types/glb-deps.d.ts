declare module 'upng-js' {
  const UPNG: {
    encode: (
      bufs: ArrayBuffer[],
      w: number,
      h: number,
      ps?: number,
      dels?: number[],
      forbidPlte?: boolean
    ) => ArrayBuffer
    decode: (buffer: ArrayBuffer) => unknown
  }
  export default UPNG
}
