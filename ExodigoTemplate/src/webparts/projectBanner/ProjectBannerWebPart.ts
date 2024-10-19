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

import * as strings from 'ProjectBannerWebPartStrings';
import ProjectBanner from './components/ProjectBanner';
import { IProjectBannerProps } from './components/IProjectBannerProps';
import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { sp } from '@pnp/sp/presets/all';


export interface IProjectBannerWebPartProps {
  description: string;
  title: string;
  SubBannerFilePicker: any;
  ButtonTitle:any;
  ButtonLink:any;
  TitleFontSize:any;
  TitleFontcolor:any;
  TitleFontAlignment:any;
  DescriptionFontSize:any;
  DescriptionFontcolor:any;
  DescriptionFontAlignment:any;
  BackgrounOverlay:any;
  BackButtonBackground:any;
  BackButtonFontSize:any;
  BackButtonFontColor:any;
  ProjectImage:boolean;
  ProjectImageFilePicker:any;
  TitleDescSpacing:any;
}

export default class ProjectBannerWebPart extends BaseClientSideWebPart<IProjectBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    sp.setup({ spfxContext: this.context });
    return super.onInit(); 
  }

  public render(): void {
    const element: React.ReactElement<IProjectBannerProps> = React.createElement(
      ProjectBanner,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        SubBannerFilePicker: this.properties.SubBannerFilePicker,
        title: this.properties.title,
        ButtonTitle:this.properties.ButtonTitle,
        ButtonLink:this.properties.ButtonLink,
        spfxContext: this.context,
        TitleFontSize:this.properties.TitleFontSize,
        TitleFontcolor:this.properties.TitleFontcolor,
        TitleFontAlignment:this.properties.TitleFontAlignment,
        DescriptionFontSize:this.properties.DescriptionFontSize,
        DescriptionFontcolor:this.properties.DescriptionFontcolor,
        DescriptionFontAlignment:this.properties.DescriptionFontAlignment,
        BackgrounOverlay:this.properties.BackgrounOverlay,
        BackButtonBackground:this.properties.BackButtonBackground,
        BackButtonFontSize:this.properties.BackButtonFontSize,
        BackButtonFontColor:this.properties.BackButtonFontColor,
        ProjectImage : this.properties.ProjectImage,
        ProjectImageFilePicker: this.properties.ProjectImageFilePicker,
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Title"
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('TitleDescSpacing', {
                  label: "Space between Title & Description"
                }),
                PropertyPaneToggle('ProjectImage', {
                  key: 'ProjectImage',
                  label: 'Show/Hide Project Image',
                  checked: true,
                  offText: "Hide Project Image",
                  onText: "Show Project Image",
                }),
                PropertyFieldFilePicker('ProjectImageFilePicker', {
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => {
                    if (e.fileAbsoluteUrl == null) {
                      e.downloadFileContent().then(async r => {
                        let fileresult = await sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/").files.add(e.fileName, r, true);
                        // this.properties.BannerFilePicker = e;
                        this.properties.ProjectImageFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
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
                        this.properties.ProjectImageFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
                        console.log(fileresult)
                        this.context.propertyPane.refresh();
                        this.render();
                      });
                    }
                   }, 
                  buttonLabel: "Project Image",
                  label: "Project Image",
                  key: 'FilePickerID',
                  filePickerResult: this.properties.ProjectImageFilePicker,
                  // hideLocalUploadTab: true,
                  hideSiteFilesTab:true, 
                  hideOneDriveTab:true,
                  hideRecentTab:true,
                  hideLinkUploadTab:true,
                  
                }),
                PropertyFieldFilePicker('SubBannerFilePicker', {
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => {
                    if (e.fileAbsoluteUrl == null) {
                      e.downloadFileContent().then(async r => {
                        let fileresult = await sp.web.getFolderByServerRelativeUrl(this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/").files.add(e.fileName, r, true);
                        // this.properties.BannerFilePicker = e;
                        this.properties.SubBannerFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
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
                        this.properties.SubBannerFilePicker = this.context.pageContext.web.serverRelativeUrl + "/SiteAssets/"+ fileresult.data.Name;
                        console.log(fileresult)
                        this.context.propertyPane.refresh();
                        this.render();
                      });
                    }
                   }, 
                  buttonLabel: "Image",
                  label: "Banner Image",
                  key: 'FilePickerID',
                  filePickerResult: this.properties.SubBannerFilePicker,
                 // hideLocalUploadTab: true,
                 hideSiteFilesTab:true, 
                 hideOneDriveTab:true,
                 hideRecentTab:true,
                 hideLinkUploadTab:true,
                }),
                PropertyPaneTextField('ButtonTitle', {
                  label: "Button Title",
                }),
                PropertyPaneTextField('ButtonLink', {
                  label: "Button Link"
                }),
                PropertyPaneTextField('BackButtonBackground', {
                  label: "Button Background"
                }),
                PropertyPaneTextField('BackButtonFontSize', {
                  label: "Button FontSize"
                }),
                PropertyPaneTextField('BackButtonFontColor', {
                  label: "Button FontColor"
                }),
                PropertyPaneTextField('TitleFontSize', {
                  label: "Title FontSize"
                }),
                PropertyPaneTextField('TitleFontcolor', {
                  label: "Title Fontcolor"
                }),
                PropertyPaneTextField('TitleFontAlignment', {
                  label: "Title FontAlignment"
                }),
                PropertyPaneTextField('DescriptionFontSize', {
                  label: "Description FontSize"
                }),
                PropertyPaneTextField('DescriptionFontcolor', {
                  label: "Description Fontcolor"
                }),
                PropertyPaneTextField('DescriptionFontAlignment', {
                  label: "Description Font Alignment"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
