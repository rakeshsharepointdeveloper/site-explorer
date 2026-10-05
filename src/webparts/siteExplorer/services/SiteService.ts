import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISiteInfo {
  id: string;
  title: string;
  url: string;
  host: string;
  path: string;
  initials: string;
  logoColor: string;
  lastModified?: Date;
  created?: Date;
  /** Same host as the page hosting the web part, so the iframe stays first-party. */
  embeddable: boolean;
}

// eslint-disable-next-line @rushstack/no-new-null -- shape of the SharePoint search REST response
interface ISearchCell { Key: string; Value: string | null; }
interface ISearchRow { Cells: ISearchCell[]; }

export interface IDocItem {
  name: string;
  serverRelativeUrl: string;
  isFolder: boolean;
  modified?: Date;
  modifiedBy?: string;
  size?: number;
  itemCount?: number;
}

export interface IRecentFile {
  name: string;
  url: string;
  location: string;
  modified?: Date;
  modifiedBy?: string;
}

export interface ISiteListInfo {
  id: string;
  title: string;
  type: string;
  isLibrary: boolean;
  itemCount: number;
  modified?: Date;
  rootUrl: string;
  viewUrl: string;
}

export interface ISiteDetails {
  description: string;
}

const LIST_TYPES: { [template: number]: string } = {
  100: 'List', 101: 'Document library', 103: 'Links', 104: 'Announcements', 106: 'Calendar',
  107: 'Tasks', 109: 'Picture library', 119: 'Site Pages', 171: 'Tasks', 851: 'Asset library'
};

interface IRawFolder { Name: string; ServerRelativeUrl: string; TimeLastModified?: string; ItemCount?: number; }
interface IRawFile { Name: string; ServerRelativeUrl: string; TimeLastModified?: string; Length?: string | number; ModifiedBy?: { Title?: string }; }

const PAGE_SIZE = 500;
const MAX_SITES = 2000;
const CONCURRENCY = 6;
const CACHE_KEY = 'siteExplorer:sites:v2';
const CACHE_TTL_MS = 10 * 60 * 1000;
// odata-version '' stops SPHttpClient sending OData-Version: 4.0, which the search endpoint can reject.
const JSON_HEADERS = { Accept: 'application/json;odata=nometadata', 'odata-version': '' };

/** Pull SharePoint's own error text out of a failed response so the UI shows the real cause. */
async function errorText(res: SPHttpClientResponse): Promise<string> {
  let detail = '';
  try {
    const body = await res.text();
    try {
      const json = JSON.parse(body);
      detail = json?.['odata.error']?.message?.value ?? json?.error?.message?.value ?? json?.error?.message ?? '';
    } catch {
      detail = body.slice(0, 200);
    }
  } catch {
    /* no body */
  }
  const id = res.headers.get('SPRequestGuid') ?? res.headers.get('request-id') ?? '';
  return `${res.status}${detail ? `: ${detail}` : ''}${id ? ` (request ${id})` : ''}`;
}

// Dark enough for white initials at 4.5:1.
const LOGO_COLORS = ['#0e5a8a', '#5c2e91', '#00666b', '#8a3707', '#3b6b0f', '#4a4fa8', '#7a2953', '#005b70'];

function initialsFor(title: string): string {
  const words = title.replace(/[^A-Za-z0-9 ]/g, ' ').split(' ').filter(w => w.length > 0);
  if (words.length === 0) return '?';
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return (words[0][0] + words[1][0]).toUpperCase();
}

function colorFor(key: string): string {
  let h = 0;
  for (let i = 0; i < key.length; i++) h = (h * 31 + key.charCodeAt(i)) | 0;
  return LOGO_COLORS[Math.abs(h) % LOGO_COLORS.length];
}

export class SiteService {
  private readonly host: string;
  private readonly libraryCache = new Map<string, string>();

  constructor(private readonly client: SPHttpClient, private readonly webUrl: string) {
    this.host = new URL(webUrl).host.toLowerCase();
  }

  public async getSites(forceRefresh = false): Promise<ISiteInfo[]> {
    if (!forceRefresh) {
      const cached = this.readCache();
      if (cached) return cached;
    }
    const sites = await this.searchSites();
    await this.enrichDates(sites);
    sites.sort((a, b) => (b.lastModified?.getTime() ?? 0) - (a.lastModified?.getTime() ?? 0));
    this.writeCache(sites);
    return sites;
  }

  /**
   * Server-relative URL of the site's main document library (the oldest visible one,
   * which is "Documents" on team sites whatever it's been renamed or translated to).
   */
  public async getLibraryRoot(site: ISiteInfo): Promise<string> {
    const cached = this.libraryCache.get(site.url);
    if (cached) return cached;
    let result = `${site.path.replace(/\/$/, '')}/Shared Documents`;
    try {
      const res = await this.client.get(
        `${site.url}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false and IsCatalog eq false` +
          `&$select=RootFolder/ServerRelativeUrl&$expand=RootFolder&$orderby=Created asc&$top=1`,
        SPHttpClient.configurations.v1,
        { headers: JSON_HEADERS }
      );
      if (res.ok) {
        const json = await res.json();
        const rel: string | undefined = json?.value?.[0]?.RootFolder?.ServerRelativeUrl;
        if (rel) result = rel;
      }
    } catch {
      /* fall back to the default name */
    }
    this.libraryCache.set(site.url, result);
    return result;
  }

  /** Absolute URL for a server-relative path on the site's host. */
  public absoluteUrl(site: ISiteInfo, serverRelativeUrl: string): string {
    return `https://${site.host}${encodeURI(serverRelativeUrl).replace(/#/g, '%23')}`;
  }

  /** Folders and files directly inside a folder, folders first, both by name. */
  public async getFolderItems(site: ISiteInfo, folderUrl: string, libraryRoot: string): Promise<IDocItem[]> {
    const path = encodeURIComponent(folderUrl.replace(/'/g, "''"));
    const base = `${site.url}/_api/web/GetFolderByServerRelativePath(decodedurl='${path}')`;
    const [folderRes, fileRes] = await Promise.all([
      this.client.get(`${base}/Folders?$select=Name,ServerRelativeUrl,TimeLastModified,ItemCount`,
        SPHttpClient.configurations.v1, { headers: JSON_HEADERS }),
      this.client.get(`${base}/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length,ModifiedBy/Title&$expand=ModifiedBy`,
        SPHttpClient.configurations.v1, { headers: JSON_HEADERS })
    ]);
    if (!folderRes.ok) throw new Error(`Couldn't open this folder (${await errorText(folderRes)}).`);
    if (!fileRes.ok) throw new Error(`Couldn't list files (${await errorText(fileRes)}).`);

    const folders: IRawFolder[] = (await folderRes.json())?.value ?? [];
    const files: IRawFile[] = (await fileRes.json())?.value ?? [];
    const formsUrl = `${libraryRoot}/Forms`.toLowerCase();
    const byName = (a: IDocItem, b: IDocItem): number => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });

    const folderItems: IDocItem[] = folders
      .filter(f => f.ServerRelativeUrl.toLowerCase() !== formsUrl)
      .map(f => ({
        name: f.Name,
        serverRelativeUrl: f.ServerRelativeUrl,
        isFolder: true,
        modified: f.TimeLastModified ? new Date(f.TimeLastModified) : undefined,
        itemCount: f.ItemCount
      }))
      .sort(byName);
    const fileItems: IDocItem[] = files
      .map(f => ({
        name: f.Name,
        serverRelativeUrl: f.ServerRelativeUrl,
        isFolder: false,
        modified: f.TimeLastModified ? new Date(f.TimeLastModified) : undefined,
        modifiedBy: f.ModifiedBy?.Title,
        size: f.Length !== undefined ? Number(f.Length) : undefined
      }))
      .sort(byName);
    return [...folderItems, ...fileItems];
  }

  /** Site description (title, dates and host come from the site list already). */
  public async getSiteDetails(site: ISiteInfo): Promise<ISiteDetails> {
    const res = await this.client.get(`${site.url}/_api/web?$select=Description`, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    if (!res.ok) throw new Error(`Couldn't load site details (${await errorText(res)}).`);
    const web = await res.json();
    return { description: web.Description ?? '' };
  }

  /** Most recently modified documents anywhere in the site, newest first (from search). */
  public async getRecentFiles(site: ISiteInfo, count = 10): Promise<IRecentFile[]> {
    const query = encodeURIComponent(`'path:"${site.url}" IsDocument:1'`);
    const select = encodeURIComponent("'Title,Filename,Path,ParentLink,LastModifiedTime,ModifiedBy,SPWebUrl'");
    const sort = encodeURIComponent("'LastModifiedTime:descending'");
    const url = `${this.webUrl}/_api/search/query?querytext=${query}&selectproperties=${select}&sortlist=${sort}&rowlimit=${count}&trimduplicates=false`;
    const res = await this.client.get(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    if (!res.ok) throw new Error(`Couldn't load recent files (${await errorText(res)}).`);
    const json = await res.json();
    const rows: ISearchRow[] = json?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];
    return rows.map(row => {
      const cell = (k: string): string => row.Cells.find(c => c.Key === k)?.Value ?? '';
      const path = cell('Path');
      const parent = cell('ParentLink');
      let location = '';
      try {
        location = decodeURIComponent(new URL(parent || path).pathname)
          .replace(site.path, '')
          .replace(/\/Forms\/.*$/, '')
          .replace(/^\//, '');
      } catch {
        /* leave blank */
      }
      const modified = cell('LastModifiedTime');
      return {
        name: cell('Filename') || cell('Title') || path,
        url: path,
        location,
        modified: modified ? new Date(modified) : undefined,
        modifiedBy: cell('ModifiedBy') || undefined
      };
    });
  }

  /** Visible lists and libraries in the site, by title. */
  public async getLists(site: ISiteInfo): Promise<ISiteListInfo[]> {
    const res = await this.client.get(
      `${site.url}/_api/web/lists?$filter=Hidden eq false and IsCatalog eq false` +
        `&$select=Id,Title,BaseTemplate,BaseType,ItemCount,LastItemUserModifiedDate,DefaultViewUrl,RootFolder/ServerRelativeUrl` +
        `&$expand=RootFolder&$orderby=Title`,
      SPHttpClient.configurations.v1,
      { headers: JSON_HEADERS }
    );
    if (!res.ok) throw new Error(`Couldn't load site contents (${await errorText(res)}).`);
    const json = await res.json();
    const lists: {
      Id: string; Title: string; BaseTemplate: number; BaseType: number; ItemCount: number;
      LastItemUserModifiedDate?: string; DefaultViewUrl?: string; RootFolder?: { ServerRelativeUrl?: string };
    }[] = json?.value ?? [];
    return lists.map(l => ({
      id: l.Id,
      title: l.Title,
      type: LIST_TYPES[l.BaseTemplate] ?? (l.BaseType === 1 ? 'Document library' : 'List'),
      isLibrary: l.BaseType === 1 && l.BaseTemplate !== 119,
      itemCount: l.ItemCount,
      modified: l.LastItemUserModifiedDate ? new Date(l.LastItemUserModifiedDate) : undefined,
      rootUrl: l.RootFolder?.ServerRelativeUrl ?? '',
      viewUrl: l.DefaultViewUrl ?? l.RootFolder?.ServerRelativeUrl ?? ''
    }));
  }

  /** Security-trimmed site collections the user can see, personal sites excluded. */
  private async searchSites(): Promise<ISiteInfo[]> {
    const byUrl = new Map<string, ISiteInfo>();
    const query = encodeURIComponent("'contentclass:STS_Site NOT WebTemplate:SPSPERS'");
    const select = encodeURIComponent("'Title,Path,SiteId,LastModifiedTime'");
    let start = 0;
    let total = Number.MAX_SAFE_INTEGER;

    while (start < total && start < MAX_SITES) {
      const url =
        `${this.webUrl}/_api/search/query?querytext=${query}&selectproperties=${select}` +
        `&rowlimit=${PAGE_SIZE}&startrow=${start}&trimduplicates=false`;
      const res: SPHttpClientResponse = await this.client.get(url, SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
      if (!res.ok) throw new Error(`Site search failed (${await errorText(res)}).`);

      const json = await res.json();
      const table = json?.PrimaryQueryResult?.RelevantResults;
      total = table?.TotalRows ?? 0;
      const rows: ISearchRow[] = table?.Table?.Rows ?? [];
      if (rows.length === 0) break;

      for (const row of rows) {
        const cell = (k: string): string => row.Cells.find(c => c.Key === k)?.Value ?? '';
        const fullUrl = cell('Path').replace(/\/$/, '');
        if (!fullUrl) continue;
        const key = fullUrl.toLowerCase();
        if (byUrl.has(key)) continue;
        const parsed = new URL(fullUrl);
        const title = cell('Title') || fullUrl;
        const searchDate = cell('LastModifiedTime');
        byUrl.set(key, {
          id: cell('SiteId') || key,
          title,
          url: fullUrl,
          host: parsed.host,
          path: parsed.pathname || '/',
          initials: initialsFor(title),
          logoColor: colorFor(key),
          lastModified: searchDate ? new Date(searchDate) : undefined,
          embeddable: parsed.host.toLowerCase() === this.host
        });
      }
      start += rows.length;
    }
    return Array.from(byUrl.values());
  }

  /**
   * Swap the search date for LastItemUserModifiedDate (ignores system changes).
   * It covers any list item, not documents only. Failures keep the search date.
   */
  private async enrichDates(sites: ISiteInfo[]): Promise<void> {
    let next = 0;
    const worker = async (): Promise<void> => {
      while (next < sites.length) {
        const site = sites[next++];
        if (!site.embeddable) continue; // other hosts aren't authenticated from this page
        try {
          const res = await this.client.get(
            `${site.url}/_api/web?$select=LastItemUserModifiedDate,Created`,
            SPHttpClient.configurations.v1,
            { headers: JSON_HEADERS }
          );
          if (!res.ok) continue;
          const web = await res.json();
          if (web.LastItemUserModifiedDate) site.lastModified = new Date(web.LastItemUserModifiedDate);
          if (web.Created) site.created = new Date(web.Created);
        } catch {
          /* keep search date */
        }
      }
    };
    await Promise.all(Array.from({ length: CONCURRENCY }, () => worker()));
  }

  private readCache(): ISiteInfo[] | undefined {
    try {
      const raw = sessionStorage.getItem(`${CACHE_KEY}:${this.host}`);
      if (!raw) return undefined;
      const parsed = JSON.parse(raw) as { at: number; sites: ISiteInfo[] };
      if (Date.now() - parsed.at > CACHE_TTL_MS) return undefined;
      return parsed.sites.map(s => ({
        ...s,
        lastModified: s.lastModified ? new Date(s.lastModified) : undefined,
        created: s.created ? new Date(s.created) : undefined
      }));
    } catch {
      return undefined;
    }
  }

  private writeCache(sites: ISiteInfo[]): void {
    try {
      sessionStorage.setItem(`${CACHE_KEY}:${this.host}`, JSON.stringify({ at: Date.now(), sites }));
    } catch {
      /* storage full or blocked */
    }
  }
}
