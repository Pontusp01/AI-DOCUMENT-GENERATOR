// src/webparts/aiDocumentGenerator/AiDocumentGeneratorWebPart.ts
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AiDocumentGeneratorWebPartStrings';
import AiDocumentGenerator from './components/AiDocumentGenerator';
import { IAiDocumentGeneratorProps } from './components/AiDocumentGenerator';
import { SharePointDocumentService } from '../../services/SharePointDocumentService';
import { getSP } from '../../utils/pnpjsConfig';

export interface IAiDocumentGeneratorWebPartProps {
  title: string;
  description: string;
}

export default class AiDocumentGeneratorWebPart extends BaseClientSideWebPart<IAiDocumentGeneratorWebPartProps> {
  // @ts-ignore - Needed for future use
  private _isDarkTheme: boolean = false;
  // @ts-ignore - Needed for future use
  private _environmentMessage: string = '';
  private _spDocService: SharePointDocumentService;

  public render(): void {
    const element: React.ReactElement<IAiDocumentGeneratorProps> = React.createElement(
      AiDocumentGenerator, 
      {
        context: this.context,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    try {
      // Initiera PnP SP med kontext
      getSP(this.context);
      
      // Initiera SharePointDocumentService
      this._spDocService = new SharePointDocumentService(this.context);
      
      // Kör connectivity test för att bestämma bästa metod för att hitta dokument
      console.log("Testar SharePoint-anslutning vid initiering...");
      await this._spDocService.testConnectivity();
      console.log("SharePoint-anslutningstest slutfört");
      
      // Skriv ut miljöinformation för felsökning
      console.log("Miljöinformation:");
      console.log(`- Web URL: ${this.context.pageContext.web.absoluteUrl}`);
      console.log(`- Site URL: ${this.context.pageContext.site.absoluteUrl}`);
      console.log(`- Användarnamn: ${this.context.pageContext.user.displayName}`);
      console.log(`- Är på lokal arbetsbänk: ${this.context.isServedFromLocalhost}`);
    }
    catch (error) {
      console.error("Fel vid initiering av webbdelen:", error);
    }

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsEnvironment);
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}