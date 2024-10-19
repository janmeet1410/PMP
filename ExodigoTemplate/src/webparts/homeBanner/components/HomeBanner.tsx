import * as React from 'react';
import styles from './HomeBanner.module.scss';
import { IHomeBannerProps } from './IHomeBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

require("../assets/css/fabric.min.css");
require("../assets/css/style.css");
// let ImageLink : string;


export default class HomeBanner extends React.Component<IHomeBannerProps, {}> {
  public render(): React.ReactElement<IHomeBannerProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    
    const ImageLink : string = this.props.BannerFilePicker == undefined ? require("../assets/Images/Picture1.jpg") :  document.location.origin + this.props.BannerFilePicker;
    const AccountImageLink : string = this.props.AccountImageFilePicker == undefined ? require("../assets/Images/Picture1.jpg") :  document.location.origin + this.props.AccountImageFilePicker;

    // const AccountImageLink = this.props.AccountImageFilePicker == undefined ? require("../assets/Images/Picture1.jpg") : this.props.AccountImageFilePicker.fileAbsoluteUrl;
    const test: string = ImageLink.replace(/ /g, "%20");
    const EncodeAccountImageLink: string = AccountImageLink.replace(/ /g, "%20");
    

    return (
      <div style={{display :"grid"}}>
        {/* <div className='Home-banner' style={{ backgroundImage: require('../assets/Images/wave1.svg') }}> */}
        <img className='Home-banner-bg' src={test}  />
        <div className='Home-banner' 
        // style={{ backgroundImage: "url("+ test +")" }}
        >
          <div style={{ backgroundColor:  this.props.BackgrounOverlay ? this.props.BackgrounOverlay : "rgb(14 44 66 / 25%)", height:'360px'}}>
            <div className='Home-banner-container Home-banner-wrapper'>
              {this.props.AccountImage == true ?
               <div className='ms-Grid-row'>
               <div className="ms-Grid-col ms-sm6 ms-md12">
                 <h1 className='Home-banner-title' style={{ fontSize : this.props.TitleFontSize ? this.props.TitleFontSize + "px" : "35px", color : this.props.TitleFontcolor ? this.props.TitleFontcolor  : "#ffffff", textAlign : this.props.TitleFontAlignment ? this.props.TitleFontAlignment  : "left"  }}>{this.props.title ? this.props.title : "Page Title"}</h1>
                 <p className='Home-banner-description' style={{ fontSize : this.props.DescriptionFontSize ? this.props.DescriptionFontSize + "px" : "17px", color : this.props.DescriptionFontcolor ? this.props.DescriptionFontcolor  : "#ffffff", textAlign : this.props.DescriptionFontAlignment ? this.props.DescriptionFontAlignment  : "left",marginTop:  this.props.TitleDescSpacing ? this.props.TitleDescSpacing + "px"   : "15px"   }}>{this.props.description ? this.props.description : "Page Description"}</p>
               </div>
             </div> 
              : 
              <div className='ms-Grid-row'>
                <div className="ms-Grid-col ms-sm6 ms-md6">
                  <h1 className='Home-banner-title' style={{ fontSize : this.props.TitleFontSize ? this.props.TitleFontSize + "px" : "35px", color : this.props.TitleFontcolor ? this.props.TitleFontcolor  : "#ffffff", textAlign : this.props.TitleFontAlignment ? this.props.TitleFontAlignment  : "left"  }}>{this.props.title ? this.props.title : "Page Title"}</h1>
                  <p className='Home-banner-description' style={{ fontSize : this.props.DescriptionFontSize ? this.props.DescriptionFontSize + "px" : "17px", color : this.props.DescriptionFontcolor ? this.props.DescriptionFontcolor  : "#ffffff", textAlign : this.props.DescriptionFontAlignment ? this.props.DescriptionFontAlignment  : "left", marginTop:  this.props.TitleDescSpacing ? this.props.TitleDescSpacing + "px"   : "15px"   }}>{this.props.description ? this.props.description : "Page Description"}</p>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6">
                  {/* {
                    this.props.AccountImage == true ? <> <img className='AccountImage' src={EncodeAccountImageLink} alt="" /> </> : ""
                  } */}
                  <img className='AccountImage' src={EncodeAccountImageLink} alt="" />
                </div>
              </div>
              }
              
            </div>
          </div>
        </div>
      </div>
    );
  } 

  public componentDidMount(): void {
    this.setState({});
  }
}
