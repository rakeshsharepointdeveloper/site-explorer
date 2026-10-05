import * as React from 'react';
import { ActionButton, DefaultButton, PrimaryButton, Icon, Spinner, SpinnerSize } from '@fluentui/react';
import styles from './SiteExplorer.module.scss';
import { ISiteInfo, SiteService } from '../services/SiteService';
import { formatDate } from './format';

export type SiteView = 'home' | 'documents' | 'contents';

const TABS: { key: SiteView; label: string }[] = [
  { key: 'home', label: 'Home' },
  { key: 'documents', label: 'Documents' },
  { key: 'contents', label: 'Site contents' }
];

export interface ISitePaneProps {
  site?: ISiteInfo;
  service: SiteService;
  /** Query string added to framed URLs, e.g. "env=Embedded". Empty adds nothing. */
  embedQuery: string;
  narrow: boolean;
  loadingSites: boolean;
  onBack: () => void;
}

const withQuery = (url: string, q: string): string => (q ? `${url}${url.indexOf('?') >= 0 ? '&' : '?'}${q}` : url);

/**
 * Right pane: the selected site's own SharePoint pages (home page first), framed so every
 * out-of-the-box feature works. In Teams desktop, SharePoint may ask the user to sign in
 * once inside the frame; that prompt comes from SharePoint, not from this web part.
 */
export const SitePane: React.FC<ISitePaneProps> = ({ site, service, embedQuery, narrow, loadingSites, onBack }) => {
  const [view, setView] = React.useState<SiteView>('home');
  const [frameUrl, setFrameUrl] = React.useState<string>();
  const [frameLoading, setFrameLoading] = React.useState(false);

  // Home page first whenever the site changes.
  React.useEffect(() => setView('home'), [site?.url]);

  // Resolve the page for the current tab.
  React.useEffect(() => {
    if (!site || !site.embeddable) {
      setFrameUrl(undefined);
      setFrameLoading(false);
      return;
    }
    let cancelled = false;
    const resolve = async (): Promise<string> => {
      if (view === 'documents') return service.absoluteUrl(site, await service.getLibraryRoot(site));
      if (view === 'contents') return `${site.url}/_layouts/15/viewlsts.aspx`;
      return site.url; // the site's home page (home.aspx or whatever its welcome page is)
    };
    setFrameLoading(true);
    resolve()
      .then(url => { if (!cancelled) setFrameUrl(withQuery(url, embedQuery)); })
      .catch(() => { if (!cancelled) setFrameUrl(withQuery(site.url, embedQuery)); });
    return () => { cancelled = true; };
  }, [site?.url, view, embedQuery]);

  if (!site) {
    return (
      <section className={styles.right}>
        <div className={styles.placeholder}>
          <Icon iconName="SplitObject" className={styles.placeholderIcon} aria-hidden="true" />
          <div className={styles.emptyTitle}>{loadingSites ? 'Your most recently updated site opens here' : 'Select a site to open it here'}</div>
          {loadingSites && <div className={styles.emptyText}>Once the list loads, the top site is selected automatically.</div>}
        </div>
      </section>
    );
  }

  return (
    <section className={styles.right} aria-label={site.title}>
      {narrow && (
        <div className={styles.backRow}>
          <ActionButton iconProps={{ iconName: 'ChevronLeft' }} text="Sites" onClick={onBack} />
        </div>
      )}

      <header className={styles.rightHeader}>
        <span className={styles.logoLarge} style={{ background: site.logoColor }} aria-hidden="true">{site.initials}</span>
        <div className={styles.headerText}>
          <h1 className={styles.siteTitle}>{site.title}</h1>
          <div className={styles.siteUrl}>{site.url}</div>
        </div>
        <DefaultButton
          className={styles.openButton}
          text={narrow ? undefined : 'Open in browser'}
          ariaLabel="Open in browser"
          iconProps={{ iconName: 'OpenInNewWindow' }}
          href={site.url}
          target="_blank"
          rel="noopener noreferrer"
        />
      </header>

      <nav className={styles.tabs} aria-label="Site pages">
        {TABS.map(t => {
          const active = site.embeddable && t.key === view;
          return (
            <button
              key={t.key}
              type="button"
              className={`${styles.tab} ${active ? styles.tabActive : ''}`}
              aria-current={active ? 'page' : undefined}
              disabled={!site.embeddable}
              onClick={() => { if (t.key !== view) setView(t.key); }}
            >
              {t.label}
            </button>
          );
        })}
      </nav>

      <div className={styles.frameArea}>
        {site.embeddable ? (
          <div className={styles.frameCard}>
            {frameLoading && <Spinner size={SpinnerSize.large} className={styles.frameSpinner} />}
            {frameUrl && (
              <iframe
                key={frameUrl}
                src={frameUrl}
                title={`${site.title} in SharePoint`}
                className={styles.frame}
                allow="clipboard-read; clipboard-write; fullscreen"
                onLoad={() => setFrameLoading(false)}
              />
            )}
          </div>
        ) : (
          <div className={`${styles.frameCard} ${styles.external}`}>
            <span className={styles.externalIcon}><Icon iconName="OpenInNewWindow" aria-hidden="true" /></span>
            <h3 className={styles.externalTitle}>This site opens in a new tab</h3>
            <p className={styles.emptyText}>
              It’s on {site.host}, a different SharePoint host from this page, so it can’t load here.
            </p>
            <PrimaryButton text="Open in browser" href={site.url} target="_blank" rel="noopener noreferrer" />
          </div>
        )}
      </div>

      <footer className={styles.footer}>
        <span>Host <strong>{site.host}</strong></span>
        <span>Created <strong>{formatDate(site.created)}</strong></span>
        <span>Last updated <strong>{formatDate(site.lastModified)}</strong></span>
      </footer>
    </section>
  );
};
