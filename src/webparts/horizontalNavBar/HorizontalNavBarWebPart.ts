import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HorizontalNavBarWebPartStrings';
import HorizontalNavBar from './components/HorizontalNavBar';
import { IHorizontalNavBarProps } from './components/IHorizontalNavBarProps';
import { getSP } from './loc/pnpjsConfig';

export interface IHorizontalNavBarWebPartProps {
  description: string;
}

export default class HorizontalNavBarWebPart extends BaseClientSideWebPart<IHorizontalNavBarWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IHorizontalNavBarProps> = React.createElement(
      HorizontalNavBar,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        
      }
    );
console.log(this.context.pageContext,'this.context.pageContext');

    ReactDom.render(element, this.domElement);
  }
   

  


  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    await super.onInit();
    getSP(this.context);
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
