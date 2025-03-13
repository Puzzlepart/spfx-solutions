/* eslint-disable prefer-spread */
import { useAnnouncementState } from './useAnnouncementState'
import { IAnnouncementProps } from './types'
import { useId } from '@fluentui/react-components'
import { useAnnouncementDataFetch } from './useProjectProvisionDataFetch'

/**
 * Component logic hook for `Announcement`. This hook is responsible for
 * fetching data, sorting, filtering and other logic.
 *
 * @param props Props
 */
export const useAnnouncement = (props: IAnnouncementProps) => {
  const { state, setState } = useAnnouncementState()
  useAnnouncementDataFetch(props, setState)

  const fluentProviderId = useId('fp-announcement')

  return {
    state,
    setState,
    fluentProviderId
  }
}
