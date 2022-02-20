import * as React from 'react';
import styles from './LocalPageNav.module.scss';
import { ILocalPageNavProps } from './ILocalPageNavProps';
import { Nav } from 'office-ui-fabric-react';

const LocalPageNav = ({ title, navLinks }: ILocalPageNavProps) => {
  return (
    <div className={styles.localPageNav}>
      <header className={styles.title}>{title}</header>
      <Nav groups={[navLinks]} />
    </div>
  );
};

export default LocalPageNav;