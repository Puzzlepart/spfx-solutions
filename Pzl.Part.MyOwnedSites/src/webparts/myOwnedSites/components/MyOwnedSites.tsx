import * as React from 'react';
import styles from './MyOwnedSites.module.scss';
import type { IMyOwnedSitesProps } from './IMyOwnedSitesProps';
import { getCreatedSites, getOwnedGroupSites } from '../helpers/data';
import { DetailsList, IColumn, IconButton, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import { ISiteResponse, ISite } from '../models/types';
import { ListColumns } from './ListColumns';
import * as strings from 'MyOwnedSitesWebPartStrings';

const { useEffect, useState } = React;

const MyOwnedSites: React.FC<IMyOwnedSitesProps> = ({ spfxContext, spClient }: IMyOwnedSitesProps) => {
  const [ownedGroupSites, setOwnedGroupSites] = useState<ISiteResponse>();
  const [ownedSites, setOwnedSites] = useState<ISiteResponse>();
  const [selectedPage, setSelectedPage] = useState<number>(1);
  const [loadFinished, setLoadFinished] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(false);

  const load = async (): Promise<void> => {
    setLoading(true);
    const nextPage = ownedGroupSites ? ownedGroupSites.nextPage : undefined;
    const groupSites = await getOwnedGroupSites(spfxContext, ownedGroupSites ? ownedGroupSites.pages : [], nextPage);
    if (!groupSites.nextPage) setLoadFinished(true);
    const sites = await getCreatedSites(spClient, spfxContext.pageContext.user.email,ownedSites ? ownedSites.pages : [], (selectedPage - 1) * 10);
    
    setOwnedSites(sites);
    setOwnedGroupSites(groupSites);
    setLoading(false);
  };

  useEffect(() => {
    if (!loadFinished && (!ownedGroupSites || (ownedGroupSites && ownedGroupSites.nextPage && ownedGroupSites.pages.length < selectedPage))) {
      //eslint-disable-next-line @typescript-eslint/no-floating-promises
      load();
    }
  }, [selectedPage]);

  const onRenderItemColumn = (item: ISite, index: number, column: IColumn): JSX.Element => {
    if (column.key === 'displayName') return <a href={item.url} target='_blank' rel='noopener noreferrer'>{item.displayName}</a>;
    return (
      <TooltipHost content={item[column.key as keyof ISite]}>
        <div>{item[column.key as keyof ISite]}</div>
      </TooltipHost>
    );
  };

  const page = ownedGroupSites ? ownedGroupSites.pages.filter(p => p.page === selectedPage)[0] : null;

  return (
    <div className={styles.myOwnedSites}>
      <>
        {loading ? <Spinner size={SpinnerSize.large} label={strings.LoadingSpinnerLabel} /> :
          <>
            {ownedGroupSites && page &&
              <>
                <DetailsList
                  items={page.sites}
                  columns={ListColumns}
                  onRenderItemColumn={onRenderItemColumn}
                />
                <div className={styles.pagination}>
                  {selectedPage > 1 &&
                    <>
                      <IconButton className={styles.toFirstPageButton} iconProps={{ iconName: 'DoubleChevronLeft' }} onClick={() => setSelectedPage(1)} />
                      <IconButton iconProps={{ iconName: 'ChevronLeft' }} onClick={() => setSelectedPage(selectedPage - 1)} />
                    </>
                  }
                  <div className={styles.pageCounter}>{selectedPage}</div>
                  {ownedGroupSites && ownedGroupSites.nextPage &&
                    <IconButton iconProps={{ iconName: 'ChevronRight' }} onClick={() => setSelectedPage(selectedPage + 1)} />
                  }
                </div>
              </>
            }
          </>
        }
      </>
    </div>
  );
}

export default MyOwnedSites;