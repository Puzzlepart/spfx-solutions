import * as React from 'react';
import styles from './MyOwnedSites.module.scss';
import type { IMyOwnedSitesProps } from './IMyOwnedSitesProps';
import { getOwnedGroupSites } from '../helpers/data';
import { DetailsList, IconButton, Spinner, SpinnerSize } from '@fluentui/react';
import { IGraphSiteResponse } from '../models/types';

const { useEffect, useState } = React;

const MyOwnedSites: React.FC<IMyOwnedSitesProps> = ({ spfxContext }: IMyOwnedSitesProps) => {
  const [ownedGroupSites, setOwnedGroupSites] = useState<IGraphSiteResponse>();
  const [selectedPage, setSelectedPage] = useState<number>(1);
  const [loadFinished, setLoadFinished] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(false);

  const load = async (): Promise<void> => {
    setLoading(true);
    const nextPage = ownedGroupSites ? ownedGroupSites.nextPage : undefined;
    const groupSites = await getOwnedGroupSites(spfxContext, ownedGroupSites ? ownedGroupSites.pages : [], nextPage);
    if (!groupSites.nextPage) setLoadFinished(true);

    setOwnedGroupSites(groupSites);
    setLoading(false);
  };

  useEffect(() => {
    if (!loadFinished && (!ownedGroupSites || (ownedGroupSites && ownedGroupSites.nextPage && ownedGroupSites.pages.length < selectedPage))) {
      //eslint-disable-next-line @typescript-eslint/no-floating-promises
      load();
    }
  }, [selectedPage]);

  const page = ownedGroupSites ? ownedGroupSites.pages.filter(p => p.page === selectedPage)[0] : null;

  return (
    <div className={styles.myOwnedSites}>
      <>
        {loading ? <Spinner size={SpinnerSize.large} label='Loading...' /> :
          <>
            {ownedGroupSites && page &&
              <>
                <DetailsList
                  items={page.sites}
                />
                <IconButton iconProps={{ iconName: 'ChevronRight' }} onClick={() => setSelectedPage(selectedPage + 1)} />
              </>
            }
          </>
        }
      </>
    </div>
  );
}

export default MyOwnedSites;