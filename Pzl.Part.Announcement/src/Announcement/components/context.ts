import { createContext, useContext } from 'react'
import { IAnnouncementProps, IAnnouncementState } from './types'

export interface IAnnouncementContext {
  props: IAnnouncementProps
  state: IAnnouncementState
  setState: (newState: Partial<IAnnouncementState>) => void
}

export const AnnouncementContext = createContext<IAnnouncementContext>(null)

/**
 * Hook to get the `AnnouncementContext`
 */
export function useAnnouncementContext() {
  return useContext(AnnouncementContext)
}
