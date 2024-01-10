import * as React from 'react';
import styles from './MyOwnedSites.module.scss';
import type { IMyOwnedSitesProps } from './IMyOwnedSitesProps';
import { getCreatedSites, getOwnedGroupSites } from '../helpers/data';
import { DetailsList, IColumn, IconButton, Pivot, PivotItem, Spinner, SpinnerSize, TooltipHost } from '@fluentui/react';
import { ISiteResponse, ISite, ISiteListPage, ResultSource } from '../models/types';
import { ListColumns } from './ListColumns';
import * as strings from 'MyOwnedSitesWebPartStrings';
import useLocationHash from './useLocationHash';

const { useEffect, useState } = React;

const MyOwnedSites: React.FC<IMyOwnedSitesProps> = ({ spfxContext, spClient, includeSPSites }: IMyOwnedSitesProps) => {
  const [ownedGroupSites, setOwnedGroupSites] = useState<ISiteResponse>();
  const [ownedSites, setOwnedSites] = useState<ISiteResponse>();
  const [selectedPage, setSelectedPage] = useState<number>(1);
  const [graphLoadFinished, setGraphLoadFinished] = useState<boolean>(false);
  const [spLoadFinished, setSPLoadFinished] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(true);
  const [selectedTab, setSelectedTab] = useState<string>(strings.GroupSitesTab);

  const hash = useLocationHash();

  const reset = (): void => {
    setOwnedGroupSites(undefined);
    setOwnedSites(undefined);
    setGraphLoadFinished(false);
    setSPLoadFinished(false);
    setSelectedPage(0);
  };

  const load = async (): Promise<void> => {
    setLoading(true);
    const nextPage = ownedGroupSites ? ownedGroupSites.nextPage : undefined;
    if (!graphLoadFinished) {
      const groupSites = await getOwnedGroupSites(spfxContext, ownedGroupSites ? ownedGroupSites.pages : [], nextPage, hash);
      setOwnedGroupSites(groupSites);
      if (!groupSites.nextPage) setGraphLoadFinished(true);
    }

    if (includeSPSites && !spLoadFinished) {
      const sites = await getCreatedSites(spClient, spfxContext.pageContext.user.email, ownedSites ? ownedSites.pages : [], (selectedPage - 1) * 10, hash);
      const loadedSitesCount = sites.pages.reduce((acc, curr) => acc + curr.sites.length, 0);
      setOwnedSites(sites);
      if (sites.totalRows && !(loadedSitesCount < sites.totalRows)) setSPLoadFinished(true);
    }

    setLoading(false);
  };

  useEffect(() => {
    if (selectedPage > 0) {
      if ((!graphLoadFinished && (!ownedGroupSites || (ownedGroupSites && ownedGroupSites.nextPage && ownedGroupSites.pages.length < selectedPage)))
        || (!spLoadFinished && (!ownedSites || (ownedSites && ownedSites.pages.length < selectedPage)))) {
        //eslint-disable-next-line @typescript-eslint/no-floating-promises
        load();
      }
    } else setSelectedPage(1);

  }, [selectedPage]);

  useEffect(() => {
    reset();
  }, [hash]);

  const onRenderItemColumn = (item: ISite, index: number, column: IColumn): JSX.Element => {
    if (column.key === 'displayName') return <a href={item.url} target='_blank' rel='noopener noreferrer'>{item.displayName}</a>;
    return (
      <TooltipHost content={item[column.key as keyof ISite]}>
        <div>{item[column.key as keyof ISite]}</div>
      </TooltipHost>
    );
  };

  const getCurrentPageContents = (): ISiteListPage | undefined => {
    if (selectedTab === strings.GroupSitesTab) {
      return ownedGroupSites ? ownedGroupSites.pages.filter(p => p.page === selectedPage)[0] : undefined;
    }
    if (selectedTab === strings.SitesTab) {
      return ownedSites ? ownedSites.pages.filter(p => p.page === selectedPage)[0] : undefined;
    }
  };

  const page = getCurrentPageContents();

  const renderList = (sites: ISiteResponse | undefined, pageContents: ISiteListPage | undefined, source: ResultSource): JSX.Element => {
    return (
      <>
        {loading ? <Spinner className={styles.loadingSpinner} size={SpinnerSize.large} label={strings.LoadingSpinnerLabel} /> :
          <>
            {sites && pageContents &&
              <>
                <DetailsList
                  items={pageContents.sites}
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
                  {sites && ((source === ResultSource.Graph && !graphLoadFinished) || (source === ResultSource.SharePoint && !spLoadFinished) || (sites.pages.length > selectedPage)) &&
                    <IconButton iconProps={{ iconName: 'ChevronRight' }} onClick={() => setSelectedPage(selectedPage + 1)} />
                  }
                </div>
              </>
            }
          </>
        }
      </>
    );
  };

  if (includeSPSites) {
    return (
      <div className={styles.myOwnedSites}>
        <>
          <Pivot onLinkClick={(tab) => {
            setSelectedTab(tab?.props.headerText || '');
            setSelectedPage(1);
          }}>
            <PivotItem headerText={strings.GroupSitesTab}>
              {renderList(ownedGroupSites, page || { page: 1, sites: [] }, ResultSource.Graph)}
            </PivotItem>
            <PivotItem headerText={strings.SitesTab}>
              {renderList(ownedSites, page || { page: 1, sites: [] }, ResultSource.SharePoint)}
            </PivotItem>
          </Pivot>
        </>
      </div>
    );
  } else return (
    <div className={styles.myOwnedSites}>
      {renderList(ownedGroupSites, page || { page: 1, sites: [] }, ResultSource.Graph)}
    </div>
  );
}

export default MyOwnedSites;