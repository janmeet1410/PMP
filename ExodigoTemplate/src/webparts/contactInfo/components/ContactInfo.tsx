import * as React from 'react';
import styles from './ContactInfo.module.scss';
import { IContactInfoProps } from './IContactInfoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';

require('../assets/css/fabric.min.css');
require('../assets/css/style.css');

export interface IContactInfoStates {
  ContactData: any;
}

export default class ContactInfo extends React.Component<IContactInfoProps, IContactInfoStates> {

  constructor(props: IContactInfoProps, state: IContactInfoStates) {
    super(props);
    this.state = {
      ContactData: []
    };
  }

  public componentDidMount = async () => {
    await this.getContactData();
  }

  private getContactData = async () => {
    try {
      const ContactData = await sp.web.lists.getByTitle(this.props.ContactTitle).items.select("Title,Phone,Email").get();
      if (ContactData.length > 0) {
        this.setState({ ContactData: ContactData });
      }
    }
    catch (error) {
      console.log(error);
    }
  }

  public render(): React.ReactElement<IContactInfoProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <div className='ContactInformation'>
      <p className='ContactInformation-Title' style={{ fontSize : this.props.ContactTitleFontSize ? this.props.ContactTitleFontSize + "px" : "20px", color : this.props.ContactTitleFontColor ? this.props.ContactTitleFontColor  : "", textAlign : this.props.ContactTitleFontAlignment ? this.props.ContactTitleFontAlignment  : "left"  }}>{this.props.ContactTitle? this.props.ContactTitle : "Client Contact Information" }</p>
      <div className='ms-Grid-row'>
      {
          this.state.ContactData.length > 0 && (
            this.state.ContactData.map((item) => { 
              // let imageURL = item.Image ? JSON.parse(item.Image).serverRelativeUrl : item.Image.Url;
              return (
                 <div className="ms-Grid-col ms-sm4 ms-md4">
                 <div className='Contact-card'>
                    {/* <img src={imageURL} /> */}
                    <div>
                      <p className='contactlabel' >Name</p>
                      <h6 style={{ fontSize : this.props.ContactPersonNameFontSize ? this.props.ContactPersonNameFontSize + "px" : "16px", color : this.props.ContactPersonNameFontColor ? this.props.ContactPersonNameFontColor  : "" }}>{item.Title ? item.Title : ''}</h6>
                      
                      <p className='contactlabel'>Phone</p>
                      <p style={{ fontSize : this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color : this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor  : "" }} className='contact-desc'>{item.Phone ? item.Phone : ''}</p>
                       
                      <p className='contactlabel'>Email</p>
                      <p style={{ fontSize : this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color : this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor  : "" }} className='contact-desc'>{item.Email ? item.Email : ''}</p>
                    </div>
                 </div>
               </div>
              );
            })
          )
        }
      </div>
   </div>
    );
  }
}
