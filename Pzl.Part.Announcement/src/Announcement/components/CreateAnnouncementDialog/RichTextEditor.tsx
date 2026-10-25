import React, { FC, useEffect, useRef } from 'react'
import {
  Toolbar,
  ToolbarButton,
  ToolbarDivider,
  Tooltip,
  useFieldControlProps_unstable
} from '@fluentui/react-components'
import {
  LinkRegular,
  TextBoldRegular,
  TextBulletListLtrRegular,
  TextItalicRegular,
  TextNumberListLtrRegular,
  TextUnderlineRegular
} from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import styles from './CreateAnnouncementDialog.module.scss'

export interface IRichTextEditorProps {
  id?: string
  /**
   * HTML
   */
  value: string
  disabled?: boolean
  placeholder?: string
  onChange: (html: string) => void
}

interface IFormatCommand {
  key: string
  label: string
  icon: JSX.Element
  command: string
}

/**
 * A small rich text editor (contentEditable + `document.execCommand`) for
 * rich text (HTML) columns. The value is the HTML SharePoint stores, so
 * announcements created here look the same as ones created in the list form.
 */
export const RichTextEditor: FC<IRichTextEditorProps> = (props) => {
  // Picks up id/aria-* from the surrounding Fluent `Field`
  const controlProps = useFieldControlProps_unstable(props, { supportsLabelFor: false })
  const editorRef = useRef<HTMLDivElement>(null)

  // Keep the DOM in sync with the value without resetting the caret while typing
  useEffect(() => {
    const editor = editorRef.current
    const value = props.value ?? ''
    if (editor && editor.innerHTML !== value) editor.innerHTML = value
  }, [props.value])

  const emitChange = () => props.onChange(editorRef.current?.innerHTML ?? '')

  const exec = (command: string, value?: string) => {
    editorRef.current?.focus()
    document.execCommand(command, false, value)
    emitChange()
  }

  const insertLink = () => {
    const url = window.prompt(strings.RichText.LinkPrompt)
    if (url) exec('createLink', url)
  }

  const commands: IFormatCommand[] = [
    { key: 'bold', label: strings.RichText.Bold, icon: <TextBoldRegular />, command: 'bold' },
    {
      key: 'italic',
      label: strings.RichText.Italic,
      icon: <TextItalicRegular />,
      command: 'italic'
    },
    {
      key: 'underline',
      label: strings.RichText.Underline,
      icon: <TextUnderlineRegular />,
      command: 'underline'
    },
    {
      key: 'ul',
      label: strings.RichText.BulletList,
      icon: <TextBulletListLtrRegular />,
      command: 'insertUnorderedList'
    },
    {
      key: 'ol',
      label: strings.RichText.NumberList,
      icon: <TextNumberListLtrRegular />,
      command: 'insertOrderedList'
    }
  ]

  return (
    <div className={[styles.richText, props.disabled ? styles.disabled : ''].join(' ')}>
      <Toolbar size='small' aria-label={strings.RichText.ToolbarLabel}>
        {commands.map((item) => (
          <Tooltip key={item.key} content={item.label} relationship='label'>
            <ToolbarButton
              icon={item.icon}
              disabled={props.disabled}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => exec(item.command)}
            />
          </Tooltip>
        ))}
        <ToolbarDivider />
        <Tooltip content={strings.RichText.Link} relationship='label'>
          <ToolbarButton
            icon={<LinkRegular />}
            disabled={props.disabled}
            onMouseDown={(event) => event.preventDefault()}
            onClick={insertLink}
          />
        </Tooltip>
      </Toolbar>
      <div
        ref={editorRef}
        id={controlProps.id}
        className={styles.editor}
        contentEditable={!props.disabled}
        role='textbox'
        aria-multiline='true'
        aria-labelledby={controlProps['aria-labelledby']}
        aria-describedby={controlProps['aria-describedby']}
        aria-invalid={controlProps['aria-invalid']}
        aria-required={controlProps['aria-required']}
        data-placeholder={props.placeholder}
        onInput={emitChange}
        onBlur={emitChange}
        suppressContentEditableWarning
      />
    </div>
  )
}
