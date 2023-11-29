import { useState } from 'react'
import { IQuickLinksState } from './types'

/**
 * Hook for QuickLinks component state
 */
export const useQuickLinksState = () => {
  const [state, $setState] = useState<IQuickLinksState>({
    linkStructure: []
  })

  const setState = (newState: Partial<IQuickLinksState>) => {
    $setState((_state) => ({ ..._state, ...newState }))
  }

  return { state, setState }
}
