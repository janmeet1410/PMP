import * as React from 'react';
import styles from './ProjectBanner.module.scss';
import { IProjectBannerProps } from './IProjectBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import $ from "jquery";
import { sp } from '@pnp/sp/presets/all';

require("../assets/css/fabric.min.css");
require("../assets/css/style.css");
let FormDigestValue;
let pal = {
  "palette" : {
    "themePrimary" : "#0e2c42",
    "themeLighterAlt" : "#d2dfe8",
    "themeLighter" : "#adc4d4",
    "themeLight" : "#8ca9bf",
    "themeTertiary" : "#6e91aa",
    "themeSecondary" : "#537a95",
    "themeDarkAlt" : "#3d6481",
    "themeDark" : "#29506c",
    "themeDarker" : "#1a3e57",
    "neutralLighterAlt" : "#faf9f8",
    "neutralLighter" : "#f3f2f1",
    "neutralLight" : "#edebe9",
    "neutralQuaternaryAlt" : "#e1dfdd",
    "neutralQuaternary" : "#d0d0d0",
    "neutralTertiaryAlt" : "#c8c6c4",
    "neutralTertiary" : "#8ca9bf",
    "neutralSecondary" : "#6e91aa",
    "neutralPrimaryAlt" : "#537a95",
    "neutralPrimary" : "#0e2c42",
    "neutralDark" : "#29506c",
    "black" : "#1a3e57",
    "white" : "#ffffff",
  }
}

export default class ProjectBanner extends React.Component<IProjectBannerProps, {}> {
  public render(): React.ReactElement<IProjectBannerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    const ImageLink : string = this.props.SubBannerFilePicker == undefined ? require("../assets/Images/Picture1.jpg") : document.location.origin + this.props.SubBannerFilePicker;
    const test: string = ImageLink.replace(/ /g, "%20");
    
    const ProjectImageLink : string = this.props.ProjectImageFilePicker == undefined ? require("../assets/Images/Picture1.jpg") :  document.location.origin + this.props.ProjectImageFilePicker;
    const EncodeProjectImageLink: string = ProjectImageLink.replace(/ /g, "%20");

    return (
      <div style={{display :"grid"}}>
        <img className='Home-banner-bg' src={test}  />
        <div className='Home-banner'
        //  style={{ backgroundImage: "url(" + test + ")" }}
        >
          <div style={{ backgroundColor: this.props.BackgrounOverlay ? this.props.BackgrounOverlay : "rgb(14 44 66 / 25%)" , height:'360px'}}>
            <div className='Home-banner-container Home-banner-wrapper'>
            {this.props.ProjectImage == true ? 
            <>
            <div className='ms-Grid-row'> 
                <div className="ms-Grid-col ms-sm12 ms-md12">
                  <a className='GoBackButton' style={{ fontSize : this.props.BackButtonFontSize ? this.props.BackButtonFontSize + "px" : "15px", color : this.props.BackButtonFontColor ? this.props.BackButtonFontColor  : "#0F5077", backgroundColor : this.props.BackButtonBackground ? this.props.BackButtonBackground  : "#ffffff"  }} href={this.props.ButtonLink ? this.props.ButtonLink : "#"}>{this.props.ButtonTitle ? this.props.ButtonTitle : "Go Back"}</a>
                  <h1 className='Home-banner-title' style={{ fontSize : this.props.TitleFontSize ? this.props.TitleFontSize + "px" : "35px", color : this.props.TitleFontcolor ? this.props.TitleFontcolor  : "#ffffff", textAlign : this.props.TitleFontAlignment ? this.props.TitleFontAlignment  : "center"  }}>{this.props.title ? this.props.title : "Page Title"}</h1>
                  <p className='Home-banner-description'  style={{ fontSize : this.props.DescriptionFontSize ? this.props.DescriptionFontSize + "px" : "17px", color : this.props.DescriptionFontcolor ? this.props.DescriptionFontcolor  : "#ffffff", textAlign : this.props.DescriptionFontAlignment ? this.props.DescriptionFontAlignment  : "center", marginTop:  this.props.TitleDescSpacing ? this.props.TitleDescSpacing + "px"   : "15px"  }}>{this.props.description ? this.props.description : "Page Description"}</p>
                </div>
              </div>
            </> 
            :
             <>
              <div className='ms-Grid-row'> 
                <div className="ms-Grid-col ms-sm6 ms-md6">
                  <a className='GoBackButton' style={{ fontSize : this.props.BackButtonFontSize ? this.props.BackButtonFontSize + "px" : "15px", color : this.props.BackButtonFontColor ? this.props.BackButtonFontColor  : "#0F5077", backgroundColor : this.props.BackButtonBackground ? this.props.BackButtonBackground  : "#ffffff"  }} href={this.props.ButtonLink ? this.props.ButtonLink : "#"}>{this.props.ButtonTitle ? this.props.ButtonTitle : "Go Back"}</a>
                  <h1 className='Home-banner-title' style={{ fontSize : this.props.TitleFontSize ? this.props.TitleFontSize + "px" : "35px", color : this.props.TitleFontcolor ? this.props.TitleFontcolor  : "#ffffff", textAlign : this.props.TitleFontAlignment ? this.props.TitleFontAlignment  : "center"  }}>{this.props.title ? this.props.title : "Page Title"}</h1>
                  <p className='Home-banner-description'  style={{ fontSize : this.props.DescriptionFontSize ? this.props.DescriptionFontSize + "px" : "17px", color : this.props.DescriptionFontcolor ? this.props.DescriptionFontcolor  : "#ffffff", textAlign : this.props.DescriptionFontAlignment ? this.props.DescriptionFontAlignment  : "center", marginTop:  this.props.TitleDescSpacing ? this.props.TitleDescSpacing + "px"  : "15px"   }}>{this.props.description ? this.props.description : "Page Description"}</p>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6">
                  <img className='AccountImage' src={EncodeProjectImageLink} alt="" />
                </div>
              </div>
             </> 
             }
            </div>
          </div>
        </div>
      </div>
    );
  }

  public async componentDidMount(): Promise<void> {
  //  this.RestRequest( this.props.spfxContext.pageContext.web.webServerRelativeUrl  +"/_api/thememanager/UpdateTenantTheme", {name:"Exodigo Theme", themeJson: JSON.stringify(pal)}); 
  await this.GetFormDigestValue();
  await this.RestRequest( this.props.spfxContext.pageContext.web.absoluteUrl  + "/_api/thememanager/ApplyTheme", {name:"Exodigo Theme", themeJson: JSON.stringify(pal)});
  await this.RestRequest( this.props.spfxContext.pageContext.web.absoluteUrl  + "/_api/thememanager/ApplyTheme", {name:"Exodigo Theme", themeJson: JSON.stringify(pal)});
  // await this.addimage(this.props.spfxContext.pageContext.web.absoluteUrl);
}


  public RestRequest(url,params) {
    var req = new XMLHttpRequest();
    req.onreadystatechange = function ()
    {
      if (req.readyState != 4) // Loaded
        return;
      console.log(req.responseText);
    };
    // Prepend web URL to url and remove duplicated slashes.
    // var webBasedUrl = (url).replace(/\/{2,}/,"/");
    req.open("POST",url,true);
    req.setRequestHeader("Content-Type", "application/json;charset=utf-8");
    req.setRequestHeader("ACCEPT", "application/json; odata.metadata=minimal");
    req.setRequestHeader("x-requestdigest", FormDigestValue);
    req.setRequestHeader("ODATA-VERSION","4.0");
    req.send(params ? JSON.stringify(params) : void 0);
  }

  public GetFormDigestValue()
  {
         // _spPageContextInfo.webAbsoluteUrl - will give absolute URL of the site where you are running the code.
         // You can replace this with other site URL where you want to apply the function
        
         $.ajax 
      ({   
          url: this.props.spfxContext.pageContext.web.absoluteUrl + "/_api/contextinfo",   
          type: "POST",   
          async: false,   
          headers: { "accept": "application/json;odata=verbose" },   
          success: function(data){   
               FormDigestValue = data.d.GetContextWebInformation.FormDigestValue; 
              console.log(FormDigestValue);            
          },   
          error: function (xhr, status, error)
          {
                console.log("Failed");
          }  
      });
  }

  public async addimage(url){

    var siteHeaderDivs = document.querySelectorAll('[data-navigationcomponent="SiteHeader"]');

    siteHeaderDivs.forEach(function(siteHeaderDiv) {
      var imgTag = document.createElement("img");
      imgTag.className = "logoImg-111";
      imgTag.alt = "";
      imgTag.setAttribute("aria-hidden", "true");
      imgTag.src =  url + "/SiteAssets/__rectSitelogo__Picture3.jpg";
      imgTag.style.minHeight = "0px";
      imgTag.style.minWidth = "0px";

      siteHeaderDiv.appendChild(imgTag);
  });

      // if (siteHeaderDivs) {
      //     var imgTag = document.createElement("img");
      //     imgTag.className = "logoImg-111";
      //     imgTag.alt = "";
      //     imgTag.setAttribute("aria-hidden", "true");
      //     imgTag.src =  url + "/SiteAssets/__rectSitelogo__Picture3.jpg";
      //     imgTag.style.minHeight = "0px";
      //     imgTag.style.minWidth = "0px";
  
      //     siteHeaderDiv.appendChild(imgTag);
      // }
  }
  
}
