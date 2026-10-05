import * as React from 'react';
import styles from './SiteExplorer.module.scss';
import { ISiteExplorerProps } from './ISiteExplorerProps';
import { ISiteInfo } from '../services/SiteService';
import { SiteList } from './SiteList';
import { SitePane } from './SitePane';

const NARROW_PX = 720;

const SiteExplorer: React.FC<ISiteExplorerProps> = ({ service, embedQuery, height }) => {
  const rootRef = React.useRef<HTMLDivElement>(null);
  const [narrow, setNarrow] = React.useState(false);
  const [sites, setSites] = React.useState<ISiteInfo[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string>();
  const [filter, setFilter] = React.useState('');
  const [selected, setSelected] = React.useState<ISiteInfo>();
  const [showSiteOnNarrow, setShowSiteOnNarrow] = React.useState(false);

  // Switch to list/detail when the web part itself is narrow (phone, Teams mobile, narrow column).
  React.useEffect(() => {
    const el = rootRef.current;
    if (!el) return;
    const update = (): void => setNarrow(el.clientWidth < NARROW_PX);
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const load = React.useCallback(async (force: boolean): Promise<void> => {
    setLoading(true);
    setError(undefined);
    try {
      const result = await service.getSites(force);
      setSites(result);
      setSelected(prev => {
        if (prev) return result.find(s => s.url === prev.url) ?? prev;
        return result.find(s => s.embeddable);
      });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  }, [service]);

  React.useEffect(() => { load(false).catch(() => undefined); }, [load]);

  const visible = React.useMemo(() => {
    const f = filter.trim().toLowerCase();
    if (!f) return sites;
    return sites.filter(s => s.title.toLowerCase().indexOf(f) >= 0 || s.url.toLowerCase().indexOf(f) >= 0);
  }, [sites, filter]);

  const onSelect = (site: ISiteInfo): void => {
    if (narrow && !site.embeddable) {
      window.open(site.url, '_blank', 'noopener,noreferrer');
      return;
    }
    setSelected(site);
    setShowSiteOnNarrow(true);
  };

  const showList = !narrow || !showSiteOnNarrow;
  const showPane = !narrow || showSiteOnNarrow;

  return (
    <div ref={rootRef} className={styles.root} style={{ height }}>
      {showList && (
        <SiteList
          sites={visible}
          totalCount={visible.length}
          loading={loading}
          error={error}
          filter={filter}
          selectedUrl={selected?.url}
          narrow={narrow}
          onFilterChange={setFilter}
          onRefresh={() => { load(true).catch(() => undefined); }}
          onSelect={onSelect}
        />
      )}
      {showPane && (
        <SitePane
          site={selected}
          service={service}
          embedQuery={embedQuery}
          narrow={narrow}
          loadingSites={loading}
          onBack={() => setShowSiteOnNarrow(false)}
        />
      )}
    </div>
  );
};

export default SiteExplorer;
