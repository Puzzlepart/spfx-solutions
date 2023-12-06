import { useState } from 'react'
import { IAllLinksState } from './types'

/**
 * Hook for AllLinks component state
 */
export const useAllLinksState = () => {
  const [state, $setState] = useState<IAllLinksState>({
    mandatoryLinks: undefined,
    editorLinks: undefined,
    categoryLinks: undefined,
    favouriteLinks: [],
    saveButtonDisabled: true,
    loading: true
  })

  const setState = (newState: Partial<IAllLinksState>) => {
    $setState((_state) => ({ ..._state, ...newState }))
  }

  return { state, setState }
}
