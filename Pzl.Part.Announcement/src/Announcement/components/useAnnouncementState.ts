/* eslint-disable prefer-spread */
import { useState } from 'react'
import { IAnnouncementState } from './types'

/**
 * Component state hook for `Announcement`.
 *
 * @param props Props
 */
export function useAnnouncementState() {
  const [state, $setState] = useState<IAnnouncementState>({
    announcements: [],
    loading: true,
    error: null
  })

  /**
   * Set state like `setState` in class components where
   * the new state is merged with the current state.
   *
   * @param newState New state
   */
  const setState = (newState: Partial<IAnnouncementState>) =>
    $setState((currentState) => ({ ...currentState, ...newState }))

  return { state, setState } as const
}
