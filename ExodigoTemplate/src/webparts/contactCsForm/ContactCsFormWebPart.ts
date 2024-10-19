import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ContactCsFormWebPartStrings';
import ContactCsForm from './components/ContactCsForm';
import { IContactCsFormProps } from './components/IContactCsFormProps';
import { sp } from '@pnp/sp/presets/all';

export interface IContactCsFormWebPartProps {
  description: string;
  FormTitle:any;
  FormTitleFontSize:any;
  FormTitleFontcolor:any;
  FormTitleFontAlignment:any;
  FormLabelFontSize:any;
  FormLabelFontColor:any;
  SubmitButtonFontSize:any;
  SubmitButtonFontColor:any;
  SubmitButtonBackground:any;
}

export default class ContactCsFormWebPart extends BaseClientSideWebPart<IContactCsFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    sp.setup({ spfxContext: this.context });

    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IContactCsFormProps> = React.createElement(
      ContactCsForm,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        spfxContext: this.context,
        FormTitle: this.properties.FormTitle,
        FormTitleFontSize: this.properties.FormTitleFontSize,
        FormTitleFontcolor: this.properties.FormTitleFontcolor,
        FormTitleFontAlignment: this.properties.FormTitleFontAlignment,
        FormLabelFontSize: this.properties.FormLabelFontSize,
        FormLabelFontColor: this.properties.FormLabelFontColor,
        SubmitButtonFontSize: this.properties.SubmitButtonFontSize,
        SubmitButtonFontColor: this.properties.SubmitButtonFontColor,
        SubmitButtonBackground: this.properties.SubmitButtonBackground,
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
                // PropertyPaneTextField('description', {
                //   label: ""
                // }),
                PropertyPaneTextField('FormTitle', {
                  label: "Form Title"
                }),
                PropertyPaneTextField('FormTitleFontSize', {
                  label: "Form Title FontSize"
                }),
                PropertyPaneTextField('FormTitleFontcolor', {
                  label: "Form Title Font color"
                }),
                PropertyPaneTextField('FormTitleFontAlignment', {
                  label: "Form Title Font Alignment"
                }),
                PropertyPaneTextField('FormLabelFontSize', {
                  label: "Form Label FontSize"
                }),
                PropertyPaneTextField('FormLabelFontColor', {
                  label: "Form Label FontColor"
                }),
                PropertyPaneTextField('SubmitButtonFontSize', {
                  label: "Submit Button FontSize"
                }),
                PropertyPaneTextField('SubmitButtonFontColor', {
                  label: "Submit Button FontColor"
                }),
                PropertyPaneTextField('SubmitButtonBackground', {
                  label: "Submit Button Background"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
