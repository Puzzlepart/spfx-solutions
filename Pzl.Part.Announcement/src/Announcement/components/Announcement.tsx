import React, { FC } from 'react'
import styles from './Announcement.module.scss'
import { IAnnouncementProps } from './types'

export const Announcement: FC<IAnnouncementProps> = (props) => {
  console.log(props)

  return (
    <div className={styles.announcement}>
      <h2>Hello world!</h2>
    </div>
  )
}
