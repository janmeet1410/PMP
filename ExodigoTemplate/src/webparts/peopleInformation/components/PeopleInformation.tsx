import * as React from 'react';
import styles from './PeopleInformation.module.scss';
import { IPeopleInformationProps } from './IPeopleInformationProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp/presets/all';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ComboBox, DefaultButton, Dialog, DialogFooter, Dropdown, IComboBoxOption, Icon, IDropdownOption, PrimaryButton, TextField } from 'office-ui-fabric-react';

require('../assets/css/fabric.min.css');
require('../assets/css/style.css');
export interface IPeopleInformationStates {
  PeopleData: any;
  ClientContactData: any;
  ExternalContactData: any;
  AddDialog1: boolean;
  AddDialog2: boolean;
  AddDialog3: boolean;
  AddName: any;
  AddPhone: any;
  AddEmail: any;
  AddRole: any;
}

const dialogContentProps = {
  title: 'Add Contact',
  closeButtonAriaLabel: 'Close'
};

const options: IDropdownOption[] = [
  { key: 'Project Manager', text: 'Project Manager' },
  { key: 'Customer Success', text: 'Customer Success' },
];

export default class PeopleInformation extends React.Component<IPeopleInformationProps, IPeopleInformationStates> {

  constructor(props: IPeopleInformationProps, state: IPeopleInformationStates) {
    super(props);
    this.state = {
      PeopleData: [],
      ClientContactData: [],
      ExternalContactData: [],
      AddDialog1: true,
      AddName: '',
      AddPhone: '',
      AddEmail: '',
      AddRole: '',
      AddDialog2: true,
      AddDialog3: true,
    };
  }

  public componentDidMount = async () => {
    await this.getPeopleData();
    await this.getClientContactData();
    await this.getExternalContactData();

  }

  private getPeopleData = async () => {
    try {
      const PeopleData = await sp.web.lists.getByTitle(this.props.ExodigoContactTitle).items.select("Title,Phone,Email,Role,ID").get();
      if (PeopleData.length > 0) {
        this.setState({ PeopleData: PeopleData });
      }
    }
    catch (error) {
      console.log(error);
    }
  }

  private getClientContactData = async () => {
    try {
      const ClientContactData = await sp.web.lists.getByTitle(this.props.ClientContactTitle).items.select("Title,Phone,Email,Role,ID").get();
      if (ClientContactData.length > 0) {
        this.setState({ ClientContactData: ClientContactData });
      }
    }
    catch (error) {
      console.log(error);
    }
  }

  private getExternalContactData = async () => {
    try {
      const ExternalContactData = await sp.web.lists.getByTitle(this.props.ManagerContactTitle).items.select("Title,Phone,Email,Role,ID").get();
      if (ExternalContactData.length > 0) {
        this.setState({ ExternalContactData: ExternalContactData });
      }
    }
    catch (error) {
      console.log(error);
    }
  }

  public addPeopleData = async () => {
    let data = {
      Title: this.state.AddName,
      Phone: this.state.AddPhone,
      Email: this.state.AddEmail,
      Role: this.state.AddRole,
    }

    await sp.web.lists
      .getByTitle('Exodigo Contacts')
      .items.add(data)
      .then((data) => {
        alert("Contact Added Successfully");
        this.getPeopleData();
        this.setState({
          AddName: "",
          AddPhone: "",
          AddEmail: "",
          AddRole: "",
          AddDialog1: true,
        });
      })
      .catch((err) => {
        console.log(err);
      });
  }

  public addClientContactData = async () => {
    let data = {
      Title: this.state.AddName,
      Phone: this.state.AddPhone,
      Email: this.state.AddEmail,
    }

    await sp.web.lists
      .getByTitle('Client Contacts')
      .items.add(data)
      .then((data) => {
        alert("Contact Added Successfully");
        this.getClientContactData();
        this.setState({
          AddName: "",
          AddPhone: "",
          AddEmail: "",
          AddDialog2: true,
        });
      })
      .catch((err) => {
        console.log(err);
      });
  }

  public addExternalContactData = async () => {
    let data = {
      Title: this.state.AddName,
      Phone: this.state.AddPhone,
      Email: this.state.AddEmail,
    }

    await sp.web.lists
      .getByTitle('Line Manager Contacts')
      .items.add(data)
      .then((data) => {
        alert("Contact Added Successfully");
        this.getExternalContactData();
        this.setState({
          AddName: "",
          AddPhone: "",
          AddEmail: "",
          AddDialog3: true,
        });
      })
      .catch((err) => {
        console.log(err);
      });
  }

  public DeletePeopleData(ID) {
    sp.web.lists.getByTitle("Exodigo Contacts").items.getById(ID).delete().then(_ => this.getPeopleData());
  }

  public DeleteClientContactData(ID) {
    sp.web.lists.getByTitle("Client Contacts").items.getById(ID).delete().then(_ => this.getClientContactData());
  }

  public DeleteExternalContactData(ID) {
    sp.web.lists.getByTitle("Line Manager Contacts").items.getById(ID).delete().then(_ => this.getExternalContactData());
  }

  public render(): React.ReactElement<IPeopleInformationProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <>
        <div className='Contact-wrapper'>

          {
            this.props.HideExodigoContacts == true ?
              <>
                <div className='PepopleInformation' style={{ flexBasis: this.props.ContactSectionPadding ? this.props.ContactSectionPadding + "%" : "33%" }}>

                  <p className='PepopleInformation-Title' style={{ fontSize: this.props.ContactTitleFontSize ? this.props.ContactTitleFontSize + "px" : "20px", color: this.props.ContactTitleFontColor ? this.props.ContactTitleFontColor : "", textAlign: this.props.ContactTitleFontAlignment ? this.props.ContactTitleFontAlignment : "left", display: 'flex' }} >{this.props.ExodigoContactTitle ? this.props.ExodigoContactTitle : "Exodigo Contacts"}<button className='Add-btn' onClick={() => this.setState({ AddDialog1: false })}>Add</button></p>
                  <div className='ms-Grid-row'>
                    {
                      this.state.PeopleData.length > 0 && (
                        this.state.PeopleData.map((item) => {
                          let imageURL = item.Image ? JSON.parse(item.Image).serverRelativeUrl : require('../assets/Images/userdefault.jpg');
                          return (
                            <div className="ms-Grid-col ms-sm12 ms-md12">
                              <div className='People-card'>
                                {/* <img src={imageURL} /> */}
                                <div>
                                  <h6 style={{ fontSize: this.props.ContactPersonNameFontSize ? this.props.ContactPersonNameFontSize + "px" : "18px", color: this.props.ContactPersonNameFontColor ? this.props.ContactPersonNameFontColor : "" }}>{item.Title ? item.Title : ''}<Icon iconName='Cancel' className='Delete-icon' onClick={() => this.DeletePeopleData(item.Id)}></Icon></h6>
                                  <p style={{ fontSize: this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color: this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor : "" }} >{item.Phone ? item.Phone : ''}</p>
                                  <a href={'mailto:' + item.Email} style={{ color: "inherit" }}>
                                    <p style={{ fontSize: this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color: this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor : "" }}>{item.Email ? item.Email : ''}</p>
                                  </a>
                                  <p style={{ fontSize: this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color: this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor : "" }}>{item.Role ? item.Role : ''}</p>
                                </div>
                              </div>
                            </div>
                          );
                        })
                      )
                    }
                  </div>
                </div>
              </> : <></>
          }

          {
            this.props.HideClientContacts == true ?
              <>
                <div className='PepopleInformation' style={{ flexBasis: this.props.ContactSectionPadding ? this.props.ContactSectionPadding + "%" : "33%" }}>
                  <p className='PepopleInformation-Title' style={{ fontSize: this.props.ContactTitleFontSize ? this.props.ContactTitleFontSize + "px" : "20px", color: this.props.ContactTitleFontColor ? this.props.ContactTitleFontColor : "", textAlign: this.props.ContactTitleFontAlignment ? this.props.ContactTitleFontAlignment : "left", display: 'flex' }} >{this.props.ClientContactTitle ? this.props.ClientContactTitle : "Client Contacts"}<button className='Add-btn' onClick={() => this.setState({ AddDialog2: false })}>Add</button></p>
                  <div className='ms-Grid-row'>
                    {
                      this.state.ClientContactData.length > 0 && (
                        this.state.ClientContactData.map((item) => {
                          let imageURL = item.Image ? JSON.parse(item.Image).serverRelativeUrl : require('../assets/Images/userdefault.jpg');
                          return (
                            <div className="ms-Grid-col ms-sm12 ms-md12">
                              <div className='People-card'>
                                {/* <img src={imageURL} /> */}
                                <div>
                                  <h6 style={{ fontSize: this.props.ContactPersonNameFontSize ? this.props.ContactPersonNameFontSize + "px" : "18px", color: this.props.ContactPersonNameFontColor ? this.props.ContactPersonNameFontColor : "" }}>{item.Title ? item.Title : ''}<Icon iconName='Cancel' className='Delete-icon' onClick={() => this.DeleteClientContactData(item.Id)}></Icon></h6>
                                  <p style={{ fontSize: this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color: this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor : "" }} >{item.Phone ? item.Phone : ''}</p>
                                  <a href={'mailto:' + item.Email} style={{ color: "inherit" }}>
                                    <p style={{ fontSize: this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color: this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor : "" }}>{item.Email ? item.Email : ''}</p>
                                  </a>
                                  {/* <p style={{ fontSize : this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color : this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor  : "" }}>{item.Role ? item.Role : ''}</p> */}
                                </div>
                              </div>
                            </div>
                          );
                        })
                      )
                    }
                  </div>
                </div>
              </> : <></>
          }


          {
            this.props.HideExternalContacts == true ?
              <>
                <div className='PepopleInformation' style={{ flexBasis: this.props.ContactSectionPadding ? this.props.ContactSectionPadding + "%" : "33%" }}>
                  <p className='PepopleInformation-Title' style={{ fontSize: this.props.ContactTitleFontSize ? this.props.ContactTitleFontSize + "px" : "20px", color: this.props.ContactTitleFontColor ? this.props.ContactTitleFontColor : "", textAlign: this.props.ContactTitleFontAlignment ? this.props.ContactTitleFontAlignment : "left", display: 'flex' }} >{this.props.ManagerContactTitle ? this.props.ManagerContactTitle : "Line Manager Contacts"}<button className='Add-btn' onClick={() => this.setState({ AddDialog3: false })}>Add</button></p>
                  <div className='ms-Grid-row'>
                    {
                      this.state.ExternalContactData.length > 0 && (
                        this.state.ExternalContactData.map((item) => {
                          let imageURL = item.Image ? JSON.parse(item.Image).serverRelativeUrl : require('../assets/Images/userdefault.jpg');
                          return (
                            <div className="ms-Grid-col ms-sm12 ms-md12">
                              <div className='People-card'>
                                {/* <img src={imageURL} /> */}
                                <div>
                                  <h6 style={{ fontSize: this.props.ContactPersonNameFontSize ? this.props.ContactPersonNameFontSize + "px" : "18px", color: this.props.ContactPersonNameFontColor ? this.props.ContactPersonNameFontColor : "" }}>{item.Title ? item.Title : ''}<Icon iconName='Cancel' className='Delete-icon' onClick={() => this.DeleteExternalContactData(item.Id)}></Icon></h6>
                                  <p style={{ fontSize: this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color: this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor : "" }} >{item.Phone ? item.Phone : ''}</p>
                                  <a href={'mailto:' + item.Email} style={{ color: "inherit" }}>
                                    <p style={{ fontSize: this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color: this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor : "" }}>{item.Email ? item.Email : ''}</p>
                                  </a>
                                  {/* <p style={{ fontSize : this.props.ContactPersonDetailFontSize ? this.props.ContactPersonDetailFontSize + "px" : "14px", color : this.props.ContactPersonDetailFontColor ? this.props.ContactPersonDetailFontColor  : "" }}>{item.Role ? item.Role : ''}</p> */}
                                </div>
                              </div>
                            </div>
                          );
                        })
                      )
                    }
                  </div>
                </div>
              </> :
              <></>
          }


        </div>


        <Dialog
          hidden={this.state.AddDialog1}
          dialogContentProps={dialogContentProps}
        >
          <div>
            <TextField label="Name" onChange={(val: any) => this.setState({ AddName: val.target.value })} />
            <TextField label="Phone" onChange={(val: any) => this.setState({ AddPhone: val.target.value })} />
            <TextField label="Email" onChange={(val: any) => this.setState({ AddEmail: val.target.value })} />
            <Dropdown
              placeholder="Select Role"
              label="Role"
              options={options}
              onChange={(event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => this.setState({ AddRole: item.key })}
            />
          </div>
          <DialogFooter>
            <PrimaryButton onClick={() => this.addPeopleData()} text="Add" />
            <DefaultButton onClick={() => this.setState({ AddDialog1: true })} text="Cancel" />
          </DialogFooter>
        </Dialog>

        {/* --------------------------------------------------------------- */}
        <Dialog
          hidden={this.state.AddDialog2}
          dialogContentProps={dialogContentProps}
        >
          <div>
            <TextField label="Name" onChange={(val: any) => this.setState({ AddName: val.target.value })} />
            <TextField label="Phone" onChange={(val: any) => this.setState({ AddPhone: val.target.value })} />
            <TextField label="Email" onChange={(val: any) => this.setState({ AddEmail: val.target.value })} />
          </div>
          <DialogFooter>
            <PrimaryButton onClick={() => this.addClientContactData()} text="Add" />
            <DefaultButton onClick={() => this.setState({ AddDialog2: true })} text="Cancel" />
          </DialogFooter>
        </Dialog>

        {/* ------------------------------------------------------------------ */}
        <Dialog
          hidden={this.state.AddDialog3}
          dialogContentProps={dialogContentProps}
        >
          <div>
            <TextField label="Name" onChange={(val: any) => this.setState({ AddName: val.target.value })} />
            <TextField label="Phone" onChange={(val: any) => this.setState({ AddPhone: val.target.value })} />
            <TextField label="Email" onChange={(val: any) => this.setState({ AddEmail: val.target.value })} />
          </div>
          <DialogFooter>
            <PrimaryButton onClick={() => this.addExternalContactData()} text="Add" />
            <DefaultButton onClick={() => this.setState({ AddDialog3: true })} text="Cancel" />
          </DialogFooter>
        </Dialog>
      </>
    );
  }
}
