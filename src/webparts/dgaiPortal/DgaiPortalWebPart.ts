import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DgaiPortalWebPartStrings';
import DgaiPortal from './components/DgaiPortal';
import { IDgaiPortalProps } from './components/IDgaiPortalProps';

import { spfi, SPFx } from '@pnp/sp';
import type { SPFI } from '@pnp/sp';

import { DgaiPortalService } from './services/DgaiPortalService';
import { DgaiHomeService } from './services/DgaiHomeService';

export interface IDgaiPortalWebPartProps {
  portalTitle: string;
  portalDescription: string;
  showQuickLinks: boolean;
  configurationListName: string;
}

export default class DgaiPortalWebPart extends BaseClientSideWebPart<IDgaiPortalWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp!: SPFI;

  public render(): void {
    const portalService = new DgaiPortalService(
      this._sp,
      this.properties.configurationListName
    );
    const homeService = new DgaiHomeService(
      this._sp,
      'DGAI_ConfiguracionInicio'
    );
    const element: React.ReactElement<IDgaiPortalProps> =
      React.createElement(
        DgaiPortal,
        {
          portalTitle: this.properties.portalTitle,
          portalDescription: this.properties.portalDescription,
          showQuickLinks: this.properties.showQuickLinks,
          portalService,
          homeService
        }
      );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    this._sp = spfi().using(SPFx(this.context));
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configuración del Portal DGAI'
          },
          groups: [
            {
              groupName: 'Contenido principal',
              groupFields: [
                PropertyPaneTextField('portalTitle', {
                  label: 'Título del portal'
                }),
                PropertyPaneTextField('portalDescription', {
                  label: 'Descripción'
                }),
                PropertyPaneToggle('showQuickLinks', {
                  label: 'Mostrar accesos rápidos',
                  onText: 'Sí',
                  offText: 'No'
                }),
                PropertyPaneTextField('configurationListName', {
                  label: 'Lista de configuración'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
