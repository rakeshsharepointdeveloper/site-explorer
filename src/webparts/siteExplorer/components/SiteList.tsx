import * as React from 'react';
import { Icon, IconButton, SearchBox, Spinner, SpinnerSize, MessageBar, MessageBarType, DefaultButton } from '@fluentui/react';
import styles from './SiteExplorer.module.scss';
import { ISiteInfo } from '../services/SiteService';
import { formatDate } from './format';

export interface ISiteListProps {
  sites: ISiteInfo[];
  totalCount: number;
  loading: boolean;
  error?: string;
  filter: string;
  selectedUrl?: string;
  narrow: boolean;
  onFilterChange: (value: string) => void;
  onRefresh: () => void;
  onSelect: (site: ISiteInfo) => void;
}

export const SiteList: React.FC<ISiteListProps> = (props) => {
  const { sites, totalCount, loading, error, filter, selectedUrl, narrow } = props;

  return (
    <aside className={`${styles.left} ${narrow ? styles.leftNarrow : ''}`} aria-label="Sites">
      <div className={styles.leftHeader}>
        <h2 className={styles.heading}>My sites</h2>
        {!loading && !error && <span className={styles.count}>{totalCount === 1 ? '1 site' : `${totalCount} sites`}</span>}
        <IconButton
          iconProps={{ iconName: 'Refresh' }}
          title="Refresh sites"
          ariaLabel="Refresh sites"
          disabled={loading}
          onClick={props.onRefresh}
        />
      </div>

      <div className={styles.filter}>
        <SearchBox
          placeholder="Filter by name or URL"
          ariaLabel="Filter sites"
          value={filter}
          onChange={(_, v) => props.onFilterChange(v ?? '')}
          onClear={() => props.onFilterChange('')}
        />
      </div>

      <div className={styles.listScroll}>
        {loading && <Spinner size={SpinnerSize.medium} label="Loading your sites" labelPosition="right" className={styles.listSpinner} />}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {error} Select refresh to try again.
          </MessageBar>
        )}

        {!loading && !error && sites.length === 0 && (
          <div className={styles.empty}>
            {filter ? (
              <>
                <div className={styles.emptyTitle}>No sites match “{filter}”</div>
                <div className={styles.emptyText}>Try part of the site name or its URL.</div>
                <DefaultButton text="Clear filter" onClick={() => props.onFilterChange('')} />
              </>
            ) : (
              <>
                <div className={styles.emptyTitle}>No sites yet</div>
                <div className={styles.emptyText}>Sites you can access appear here once search has indexed them.</div>
              </>
            )}
          </div>
        )}

        {!loading && (
          <ul className={styles.list}>
            {sites.map(site => {
              const active = site.url === selectedUrl && !narrow;
              return (
                <li key={site.url}>
                  <button
                    type="button"
                    className={`${styles.item} ${active ? styles.itemActive : ''}`}
                    aria-current={active ? 'true' : undefined}
                    title={site.url}
                    onClick={() => props.onSelect(site)}
                  >
                    <span className={styles.logo} style={{ background: site.logoColor }} aria-hidden="true">{site.initials}</span>
                    <span className={styles.itemText}>
                      <span className={styles.itemTitle}>{site.title}</span>
                      <span className={styles.itemMeta}>{site.path}</span>
                      <span className={styles.itemMeta}>Updated {formatDate(site.lastModified)}</span>
                      {!site.embeddable && (
                        <span className={styles.itemExternal}>
                          <Icon iconName="OpenInNewWindow" aria-hidden="true" /> Opens in a new tab
                        </span>
                      )}
                    </span>
                    {narrow && <Icon iconName={site.embeddable ? 'ChevronRight' : 'OpenInNewWindow'} className={styles.chevron} aria-hidden="true" />}
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </div>
    </aside>
  );
};
