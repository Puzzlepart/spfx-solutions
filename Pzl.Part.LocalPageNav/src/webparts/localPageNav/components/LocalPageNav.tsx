import * as React from 'react';
import styles from './LocalPageNav.module.scss';
import { INavLinkGroup, Nav } from 'office-ui-fabric-react';

export interface ILocalPageNavProps {
  title: string,
  navLinks: INavLinkGroup,
}

export const LocalPageNav: React.FunctionComponent<ILocalPageNavProps> = ({ title, navLinks }: ILocalPageNavProps) => {
  return (
    <div className={styles.localPageNav}>
      <header className={styles.title}>{title}</header>
      <Nav groups={[navLinks]} />
    </div>
  );
};

export default LocalPageNav;