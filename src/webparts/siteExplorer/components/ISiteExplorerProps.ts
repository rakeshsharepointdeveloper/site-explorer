import { SiteService } from '../services/SiteService';

export interface ISiteExplorerProps {
  service: SiteService;
  /** Query string added to framed URLs to trim SharePoint chrome, e.g. "env=Embedded". Empty disables it. */
  embedQuery: string;
  /** CSS height of the control; the Teams tab passes "100vh". */
  height: string;
}
