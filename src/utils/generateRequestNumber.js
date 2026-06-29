let seqCounter = 0

export function generateRequestNumber() {
  const year = new Date().getFullYear().toString().slice(-2)
  seqCounter = (seqCounter % 999999) + 1
  return `WE${year}-${String(seqCounter).padStart(6, '0')}`
}
