import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SiteExplorerWebPartStrings';
import SiteExplorer from './components/SiteExplorer';
import { ISiteExplorerProps } from './components/ISiteExplorerProps';
import { SiteService } from './services/SiteService';

export interface ISiteExplorerWebPartProps {
  embedQuery: string;
}

export default class SiteExplorerWebPart extends BaseClientSideWebPart<ISiteExplorerWebPartProps> {
  private _service!: SiteService;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._service = new SiteService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    // Teams host: no page margins, the web part owns the whole tab.
    document.documentElement.style.height = '100%';
    document.body.style.margin = '0';
  }

  public render(): void {
    const element: React.ReactElement<ISiteExplorerProps> = React.createElement(SiteExplorer, {
      service: this._service,
      embedQuery: (this.properties.embedQuery ?? '').trim(),
      height: '100vh'
    });
    ReactDom.render(element, this.domElement);
  }

  /** Teams light, dark and high-contrast themes arrive here; map them onto the CSS variables the styles use. */
  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    if (!theme) return;
    const s = theme.semanticColors ?? {};
    const p = theme.palette ?? {};
    const dark = !!theme.isInverted;
    const vars: Record<string, string | undefined> = {
      '--se-text': s.bodyText,
      '--se-subtle': s.bodySubtext,
      '--se-surface': s.bodyBackground,
      '--se-pane': p.neutralLighterAlt,
      '--se-well': p.neutralLighter,
      '--se-row': p.neutralLighter,
      '--se-selected': p.neutralLight,
      '--se-border': p.neutralLight,
      '--se-accent': p.themePrimary,
      '--se-accent-soft': p.themeLighter,
      '--se-disabled': p.neutralTertiary,
      '--se-warn': dark ? '#f4bfab' : '#8a3707'
    };
    Object.keys(vars).forEach(k => {
      const v = vars[k];
      if (v) this.domElement.style.setProperty(k, v);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Only reachable when a channel tab is added or its settings are edited (personal apps have no settings).
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneTextField('embedQuery', {
                  label: strings.EmbedQueryFieldLabel,
                  description: strings.EmbedQueryFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
