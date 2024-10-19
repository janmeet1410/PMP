import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ContactInfoWebPartStrings';
import ContactInfo from './components/ContactInfo';
import { IContactInfoProps } from './components/IContactInfoProps';
import { sp } from '@pnp/sp/presets/all';


export interface IContactInfoWebPartProps {
  description: string;
  ContactTitle:string;
  ContactTitleFontSize:any; 
  ContactTitleFontColor:any; 
  ContactTitleFontAlignment:any; 
  ContactPersonNameFontSize:any; 
  ContactPersonNameFontColor:any; 
  ContactPersonDetailFontSize:any; 
  ContactPersonDetailFontColor:any; 
}

export default class ContactInfoWebPart extends BaseClientSideWebPart<IContactInfoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    sp.setup({ spfxContext: this.context });
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IContactInfoProps> = React.createElement(
      ContactInfo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        spfxContext: this.context,
        ContactTitle: this.properties.ContactTitle,
        ContactTitleFontSize: this.properties.ContactTitleFontSize,
        ContactTitleFontColor: this.properties.ContactTitleFontColor,
        ContactTitleFontAlignment: this.properties.ContactTitleFontAlignment,
        ContactPersonNameFontSize: this.properties.ContactPersonNameFontSize,
        ContactPersonNameFontColor: this.properties.ContactPersonNameFontColor,
        ContactPersonDetailFontSize: this.properties.ContactPersonDetailFontSize,
        ContactPersonDetailFontColor: this.properties.ContactPersonDetailFontColor,
      }
    );

    ReactDom.render(element, this.domElement);
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

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
                }),
                PropertyPaneTextField('ContactTitle', {
                  label: "Contact Title"
                }),
                PropertyPaneTextField('ContactTitleFontSize', {
                  label: "Contact Title FontSize"
                }),
                PropertyPaneTextField('ContactTitleFontColor', {
                  label: "Contact Title FontColor"
                }),
                PropertyPaneTextField('ContactTitleFontAlignment', {
                  label: "Contact Title Font Alignment"
                }),
                PropertyPaneTextField('ContactPersonNameFontSize', {
                  label: "Contact Person Name FontSize"
                }),
                PropertyPaneTextField('ContactPersonNameFontColor', {
                  label: "Contact Person Name FontColor"
                }),
                PropertyPaneTextField('ContactPersonDetailFontSize', {
                  label: "Contact Person Detail FontSize"
                }),
                PropertyPaneTextField('ContactPersonDetailFontColor', {
                  label: "Contact Person Detail FontColor"
                }),
                
              ]
            }
          ]
        }
      ]
    };
  }
}
