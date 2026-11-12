import {
  FluentProvider,
  IdPrefixProvider,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  useId,
  webLightTheme
} from '@fluentui/react-components'
import { CalendarClock20Regular, Pin16Filled } from '@fluentui/react-icons'
import React, { FC } from 'react'
import ReactMarkdown from 'react-markdown'
import rehypeRaw from 'rehype-raw'
import { IUserMessageProps } from './types'
import { useUserMessage } from './useUserMessage'
import styles from './UserMessage.module.scss'
import { StatusBadge } from '../StatusBadge'

/**
 * A component that supports a MessageBar with markdown using react-markdown
 *
 * @category UserMessage
 */
export const UserMessage: FC<IUserMessageProps> = (props: IUserMessageProps) => {
  const fluentProviderId = useId('fp-announcement-user-message')
  const messageProps = useUserMessage(props)

  return (
    <IdPrefixProvider value={fluentProviderId}>
      <FluentProvider
        theme={webLightTheme}
        className={[props.className, styles.userMessage].filter(Boolean).join(' ')}
        style={props.containerStyle}
        hidden={props.hidden}
        onClick={props.onClick}>
        <MessageBar
          {...messageProps}
          className={styles.message}
          intent={props.intent}
          icon={props.status === 'planned' ? <CalendarClock20Regular /> : undefined}>
          <MessageBarBody>
            {props.title && (
              <MessageBarTitle>
                {props.pinned && (
                  <Pin16Filled
                    className={styles.pin}
                    title={props.pinnedTitle}
                    aria-label={props.pinnedTitle}
                  />
                )}
                {props.title}
                {props.status && <StatusBadge status={props.status} className={styles.status} />}
              </MessageBarTitle>
            )}
            {props.text && (
              <ReactMarkdown linkTarget={props.linkTarget} rehypePlugins={[rehypeRaw]}>
                {props.text}
              </ReactMarkdown>
            )}
            {props.children && props.children}
          </MessageBarBody>
        </MessageBar>
      </FluentProvider>
    </IdPrefixProvider>
  )
}

UserMessage.defaultProps = {
  linkTarget: '_blank',
  style: {}
}

export * from './types'
