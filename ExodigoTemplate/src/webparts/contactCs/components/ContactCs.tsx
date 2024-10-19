import * as React from 'react';
import styles from './ContactCs.module.scss';
import { IContactCsProps } from './IContactCsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton } from 'office-ui-fabric-react';

require("../assets/style.css");

export default class ContactCs extends React.Component<IContactCsProps, {}> {
  public render(): React.ReactElement<IContactCsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <>
      <div style={{textAlign : this.props.ContactButtonAlignment ? this.props.ContactButtonAlignment : "Center" }}>
        <a data-interception="off" target='_blank' href={this.props.ContactButtonLink ? this.props.ContactButtonLink  : this.props.spfxContext.pageContext.web.serverRelativeUrl + "/SitePages/Form.aspx"} className='ContactUs-btn' style={{fontSize: this.props.ContactButtonFontSize ? this.props.ContactButtonFontSize + "px" : "16px",color: this.props.ContactButtonFontColor ? this.props.ContactButtonFontColor: "#ffffff", backgroundColor : this.props.ContactButtonBackground ? this.props.ContactButtonBackground: "#0e2c42"}}>{this.props.ContactButtonTitle ? this.props.ContactButtonTitle  : "Contact Us"}</a>
      </div>
      </>
    );
  }
}
