import * as React from 'react';
import styles from './ContactCsForm.module.scss';
import { IContactCsFormProps } from './IContactCsFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp/presets/all';
import { IHttpClientOptions, HttpClientResponse, HttpClient } from '@microsoft/sp-http';
require('../assets/css/fabric.min.css');
require('../assets/css/style.css');

export interface IContactCsFormStates {
  CSEmails: any;
  FirstName:string;
  LastName:string;
  Email:string;
  PhoneNumber:any;
  CompanyName:string;
  MsgDetail:string;
}

export default class ContactCsForm extends React.Component<IContactCsFormProps, IContactCsFormStates> {

  constructor(props: IContactCsFormProps, state: IContactCsFormStates) {
    super(props);
    this.state = {
      CSEmails: [],
      FirstName:'',
      LastName:'',
      Email:'',
      PhoneNumber: '0',
      CompanyName:'',
      MsgDetail:'',
    };
  }

  public componentDidMount = async () => {
    await this.getCSEmailData();
  }

  private getCSEmailData = async () => {
    try {
      let AllEmails = [];
      const PeopleData = await sp.web.lists.getByTitle('Contact Information').items.select("Email,Role").filter("Role eq 'CS'").get();
      // if (PeopleData.length > 0) {
      //   this.setState({ CSEmails: PeopleData });
        console.log(PeopleData);
      // }
      if (PeopleData.length > 0) {
        PeopleData.forEach((item) => {
          AllEmails.push({
            Email: item.Email ? item.Email : '',
          });
        });
        this.setState({ CSEmails: AllEmails });
        console.log(this.state.CSEmails);
      }
    }
    catch (error) {
      console.log(error);
    }
  }

  public render(): React.ReactElement<IContactCsFormProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
     <div className='ms-Grid ContactCSForm'>
      <p className='ContactCS-title' style={{fontSize: this.props.FormTitleFontSize ? this.props.FormTitleFontSize + "px" : "20px",color: this.props.FormTitleFontcolor ? this.props.FormTitleFontcolor: "rgb(14, 44, 66)" , textAlign : this.props.FormTitleFontAlignment ? this.props.FormTitleFontAlignment: "left"}} > {this.props.FormTitle ? this.props.FormTitle: "Contact Us"}</p>
        <div className='ms-Grid-row'>
          <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mb-10'>
            <Label style={{fontSize: this.props.FormLabelFontSize ? this.props.FormLabelFontSize + "px" : "14px",color: this.props.FormLabelFontColor ? this.props.FormLabelFontColor: "rgb(14, 44, 66)"}}>First Name</Label>
            <TextField onChange={(val: any) => { this.setState({ FirstName: val.target["value"] }) }} />
          </div>
          <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mb-10'>
            <Label style={{fontSize: this.props.FormLabelFontSize ? this.props.FormLabelFontSize + "px" : "14px",color: this.props.FormLabelFontColor ? this.props.FormLabelFontColor: "rgb(14, 44, 66)"}}>Last Name</Label>
            <TextField onChange={(val: any) => { this.setState({ LastName: val.target["value"] }) }} />
          </div>
          <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 mb-10'>
            <Label style={{fontSize: this.props.FormLabelFontSize ? this.props.FormLabelFontSize + "px" : "14px",color: this.props.FormLabelFontColor ? this.props.FormLabelFontColor: "rgb(14, 44, 66)"}}>Email</Label>
            <TextField  onChange={(val: any) => { this.setState({ Email: val.target["value"] }) }} />
          </div>
          <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mb-10'>
            <Label style={{fontSize: this.props.FormLabelFontSize ? this.props.FormLabelFontSize + "px" : "14px",color: this.props.FormLabelFontColor ? this.props.FormLabelFontColor: "rgb(14, 44, 66)"}}>Phone Number</Label>
            <TextField onChange={(val: any) => { this.setState({ PhoneNumber: val.target["value"] }) }} />
          </div>
          <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 mb-10'>
            <Label style={{fontSize: this.props.FormLabelFontSize ? this.props.FormLabelFontSize + "px" : "14px",color: this.props.FormLabelFontColor ? this.props.FormLabelFontColor: "rgb(14, 44, 66)"}}>Company/Agency Name</Label>
            <TextField  onChange={(val: any) => { this.setState({ CompanyName: val.target["value"] }) }} />
          </div>
          <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 mb-10'>
            <Label style={{fontSize: this.props.FormLabelFontSize ? this.props.FormLabelFontSize + "px" : "14px",color: this.props.FormLabelFontColor ? this.props.FormLabelFontColor: "rgb(14, 44, 66)"}}>How can we help you?</Label>
            <TextField  multiline onChange={(val: any) => { this.setState({ MsgDetail: val.target["value"] }) }} />
          </div>
          <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 mb-10' style={{textAlign:'center',marginTop:'15px'}}>
            <PrimaryButton style={{fontSize: this.props.SubmitButtonFontSize ? this.props.SubmitButtonFontSize + "px" : "14px",color: this.props.SubmitButtonFontColor ? this.props.SubmitButtonFontColor: "#ffffff", backgroundColor : this.props.SubmitButtonBackground ? this.props.SubmitButtonBackground: "rgb(14, 44, 66)"}} text="Submit" onClick={()=> this.sendEmail()}></PrimaryButton>
          </div>
        </div>
     </div>
    );
  }

  private sendEmail(): Promise<HttpClientResponse> {

    if(this.state.FirstName != "" || this.state.LastName != "" || this.state.PhoneNumber != "0" || this.state.Email != "" || this.state.CompanyName != "" || this.state.MsgDetail != "" ){
      const postURL = "https://prod2-35.southeastasia.logic.azure.com:443/workflows/84f961528e864103955be62315e3e29e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=WNUPDx7YWqDK9EWRtJJqQ9Am9RDcnCStCUhzrBym4pQ";
    const body: string = JSON.stringify({
      'CSEmails' :this.state.CSEmails,
      'First Name': this.state.FirstName,
      'Last Name': this.state.LastName,
      'Email': this.state.Email,
      'Phone Number': this.state.PhoneNumber,
      'Company Name': this.state.CompanyName,
      'How we can help you ?': this.state.MsgDetail,

    });
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    const httpClientOptions: IHttpClientOptions = {
      body: body,
      headers: requestHeaders
    };
    console.log("Sending Email");

    
    return this.props.spfxContext.httpClient.post(postURL, HttpClient.configurations.v1, httpClientOptions).then((response) => {
      console.log("Flow Triggered Successfully...");
      alert('Form Details submitted Successfully.')
      this.setState({FirstName : "", LastName : "",Email : "", PhoneNumber : 0, CompanyName : "", MsgDetail : "" })
    }).catch(error => {
        console.log(error);
    });
    }
    else{
      alert('Please fill all fields.')
    }

   

  }
}
