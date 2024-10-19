import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HomeBannerWebPartStrings';
import HomeBanner from './components/HomeBanner';
import { IHomeBannerProps } from './components/IHomeBannerProps';
import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { sp } from '@pnp/sp/presets/all';

export interface IHomeBannerWebPartProps {
  description: string;
  title: string;
  BannerFilePicker: any;
  AccountImage:boolean;
  AccountImageFilePicker:any;
  TitleFontSize:any;
  TitleFontcolor:any;
  TitleFontAlignment:any;
  DescriptionFontSize:any;
  DescriptionFontcolor:any;
  DescriptionFontAlignment:any;
  BackgrounOverlay:any;
  TitleDescSpacing:any;

}

export default class HomeBannerWebPart extends BaseClientSideWebPart<IHomeBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    sp.setup({ spfxContext: this.context });

    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IHomeBannerProps> = React.createElement(
      HomeBanner,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        BannerFilePicker: this.properties.BannerFilePicker,
        title: this.properties.title,
        AccountImage : this.properties.AccountImage,
        AccountImageFilePicker: this.properties.AccountImageFilePicker,
        spfxContext: this.context,
        TitleFontSize: this.properties.TitleFontSize,
        TitleFontcolor: this.properties.TitleFontcolor,
        TitleFontAlignment: this.properties.TitleFontAlignment,
        DescriptionFontSize: this.properties.DescriptionFontSize,
        DescriptionFontcolor: this.properties.DescriptionFontcolor,
        DescriptionFontAlignment: this.properties.DescriptionFontAlignment,
        BackgrounOverlay: this.properties.BackgrounOverlay,
        TitleDescSpacing:this.properties.TitleDescSpacing,

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
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Title"
                }),
                PropertyPaneTextField('description', {
                  label: "Description",
                  multiline: true
                }),
                PropertyPaneTextField('TitleDescSpacing', {
                  label: "Space between Title & Description"
                }),
                PropertyFieldFilePicker('BannerFilePicker', {
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => {
                    if (e.fileAbsoluteUrl == null) {
                      e.downloadFileContent().then(async r => {
                        let fileresult = await sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/").files.add(e.fileName, r, true);
                        // this.properties.BannerFilePicker = e;
                        this.properties.BannerFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
                        console.log(fileresult)
                        this.context.propertyPane.refresh();
                        this.render();
                      });
                    }
                   },
                  onChanged: (e: IFilePickerResult) => { 
                    if (e.fileAbsoluteUrl == null) {
                      e.downloadFileContent().then(async r => {
                        let fileresult = await sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/").files.add(e.fileName, r, true);
                        // this.properties.BannerFilePicker = e;
                        this.properties.BannerFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
                        console.log(fileresult)
                        this.context.propertyPane.refresh();
                        this.render();
                      });
                    }
                   },
                  buttonLabel: "Image",
                  label: "Banner Image",
                  key: 'FilePickerID',
                  filePickerResult: this.properties.BannerFilePicker,
                  // hideLocalUploadTab: true,
                  hideSiteFilesTab:true, 
                  hideOneDriveTab:true,
                  hideRecentTab:true,
                  hideLinkUploadTab:true
                }),
                PropertyPaneToggle('AccountImage', {
                  key: 'AccountImage',
                  label: 'Show/Hide Account Image',
                  checked: true,
                  offText: "Hide Account Image",
                  onText: "Show Account Image",
                }),
                PropertyFieldFilePicker('AccountImageFilePicker', {
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => {
                    if (e.fileAbsoluteUrl == null) {
                      e.downloadFileContent().then(async r => {
                        let fileresult = await sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/").files.add(e.fileName, r, true);
                        // this.properties.BannerFilePicker = e;
                        this.properties.AccountImageFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
                        console.log(fileresult)
                        this.context.propertyPane.refresh();
                        this.render();
                      });
                    }
                   },
                  onChanged: (e: IFilePickerResult) => { 
                    if (e.fileAbsoluteUrl == null) {
                      e.downloadFileContent().then(async r => {
                        let fileresult = await sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/").files.add(e.fileName, r, true);
                        // this.properties.BannerFilePicker = e;
                        this.properties.AccountImageFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
                        console.log(fileresult)
                        this.context.propertyPane.refresh();
                        this.render();
                      });
                    }
                   }, 
                  buttonLabel: "Account Image",
                  label: "Account Image",
                  key: 'FilePickerID',
                  filePickerResult: this.properties.AccountImageFilePicker,
                  // hideLocalUploadTab: true,
                  hideSiteFilesTab:true, 
                  hideOneDriveTab:true,
                  hideRecentTab:true,
                  hideLinkUploadTab:true,
                  
                }),
                PropertyPaneTextField('TitleFontSize', {
                  label: "Title FontSize"
                }),
                PropertyPaneTextField('TitleFontcolor', {
                  label: "Title Fontcolor"
                }),
                PropertyPaneTextField('TitleFontAlignment', {
                  label: "Title Alignment"
                }),
                PropertyPaneTextField('DescriptionFontSize', {
                  label: "Description FontSize"
                }),
                PropertyPaneTextField('DescriptionFontcolor', {
                  label: "Description Fontcolor"
                }),
                PropertyPaneTextField('DescriptionFontAlignment', {
                  label: "Description Alignment"
                }),
                PropertyPaneTextField('BackgrounOverlay', {
                  label: "Backgroun Overlay"
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  // private saveIntoSharePoint = async(file: IFilePickerResult) => {
  //   if (file.fileAbsoluteUrl == null) {
  //     file.downloadFileContent()
  //       .then(async r => {
  //         // let fileresult = await sp.web.getFolderByServerRelativeUrl(`/sites/${this.context.pageContext.site}/SiteAssets`).files.add(file.fileName, r, true);
  //         let fileresult = await sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/").files.add(file.fileName, r, true);
  //         // this.setState({ iconProperty: document.location.origin + fileresult.data.ServerRelativeUrl });
  //         this.properties.BannerFilePicker = document.location.origin + fileresult.data.ServerRelativeUrl
  //         console.log(fileresult)
  //       });
  //   }
  //   // else {
  //   //   this.setState({ iconProperty: file.fileAbsoluteUrl });
  //   // }
  // }
}
