import React, { Component } from 'react';
import './App.css';

import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';

class App extends Component {


  office = window.Office;
  officeMailBoxItem = window.Office.context.mailbox.item;
 
  constructor(props) {
    super(props);
    this.setSubject("Leave Request");
  }


  setSubject = (subject)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    this.setState({ subject: subject})
    this.officeMailBoxItem.subject.setAsync(subject);
    
  };

   setToday = ()=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
     let today = new Date();
     this.setState({ startDateTime: today })
     this.officeMailBoxItem.start.setAsync(today );
     this.officeMailBoxItem.end.setAsync(today);
  };


   setStartdDate = (value)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    // this.setState({ value: date });
     let date = new Date(value);
     this.setState({ startDateTime: date })
     this.officeMailBoxItem.start.setAsync(date);
  };

  setEndDate = (value)=>{
    let date = new Date(value);
    this.setState({ endDateTime: date })
    this.officeMailBoxItem.end.setAsync(date);
  };

  setLeaveType = (option) => {
    this.setState({ leaveType: option })
  };

  setReason = (value) => {
    this.setState({ reason : value })
  };

  setMessage = ()=>{
    this.officeMailBoxItem.body.setAsync(
      '<div>Leave Type :' + this.state.leaveType.text + '</div>' +
      '<div>Reason :' + this.state.reason + '</div>',

      { coercionType: this.office.CoercionType.Html })
  };
  
  render() {
    return (
      <Fabric>
            <TextField label='Reason'  onChanged={this.setReason} /> 
            <DatePicker label='Start Date' onSelectDate={this.setStartdDate} />
            <DatePicker label='End Date' onSelectDate={this.setEndDate} />
            <Dropdown label='Leave Type' onChanged={this.setLeaveType} options={[
              {key:'annual',text:'Annual Leave'},
              {key:'personal',text:'Sick/Carer Leave'}
            ]} > </Dropdown>
            <DefaultButton primary={ true } onClick={this.setToday} >
              Set Today
            </DefaultButton>

        <DefaultButton primary={true} onClick={this.setMessage} >
               Prepare Message
            </DefaultButton>

      </Fabric>
    );
  }
}

export default App;