import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ContactCsWebPartStrings';
import ContactCs from './components/ContactCs';
import { IContactCsProps } from './components/IContactCsProps';
import { sp } from '@pnp/sp/presets/all';

export interface IContactCsWebPartProps {
  description: string;
  ContactButtonTitle:any;
  ContactButtonLink:any;
  ContactButtonFontSize:any;
  ContactButtonFontColor:any;
  ContactButtonBackground:any;
  ContactButtonAlignment:any;
}

export default class ContactCsWebPart extends BaseClientSideWebPart<IContactCsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    sp.setup({ spfxContext: this.context });

    this._environmentMessage = this._getEnvironmentMessage();
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IContactCsProps> = React.createElement(
      ContactCs,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        ContactButtonTitle: this.properties.ContactButtonTitle,
        ContactButtonLink: this.properties.ContactButtonLink,
        ContactButtonFontSize: this.properties.ContactButtonFontSize,
        ContactButtonFontColor: this.properties.ContactButtonFontColor,
        ContactButtonBackground: this.properties.ContactButtonBackground,
        ContactButtonAlignment: this.properties.ContactButtonAlignment,
        spfxContext: this.context,

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
                PropertyPaneTextField('ContactButtonTitle', {
                  label: "Contact Button Title"
                }),
                PropertyPaneTextField('ContactButtonLink', {
                  label: "Contact Button Link"
                }),
                PropertyPaneTextField('ContactButtonFontSize', {
                  label: "Contact Button FontSize"
                }),
                PropertyPaneTextField('ContactButtonFontColor', {
                  label: "Contact Button FontColor"
                }),
                PropertyPaneTextField('ContactButtonBackground', {
                  label: "Contact Button Background"
                }),
                PropertyPaneTextField('ContactButtonAlignment', {
                  label: "Contact Button Alignment"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
