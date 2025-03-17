import { IUserMessageProps } from './types'
import { CSSProperties } from 'react'

export function useUserMessage(props: IUserMessageProps): { styles: CSSProperties } {
  let styles: CSSProperties = {}

  if (props.fixedCenter) {
    styles = {
      ...(styles['root'] as any),
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      minHeight: props.fixedCenter
    }
  }

  if (props.isCompact) {
    styles = {
      marginTop: '3px',
      marginBottom: '0px'
    }
  }
  return { styles } as const
}
