import { useEffect, useCallback } from 'react'

const STORAGE_KEY = 'we-registration-draft'

export function useAutoSave(data) {
  useEffect(() => {
    if (!data || Object.keys(data).length === 0) return
    const timer = setTimeout(() => {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data))
    }, 1000)
    return () => clearTimeout(timer)
  }, [data])
}

export function getSavedDraft() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY)
    return saved ? JSON.parse(saved) : null
  } catch {
    return null
  }
}

export function clearSavedDraft() {
  localStorage.removeItem(STORAGE_KEY)
}
