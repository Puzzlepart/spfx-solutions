import * as React from 'react';
import styles from './MyOwnedSites.module.scss';
import type { IMyOwnedSitesProps } from './IMyOwnedSitesProps';
import { getOwnedGroupSites } from '../helpers/data';
import { DetailsList, IColumn, IconButton, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import { IGraphSiteResponse, ISite } from '../models/types';
import { ListColumns } from './ListColumns';

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
        {loading ? <Spinner size={SpinnerSize.large} label='Loading...' /> :
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