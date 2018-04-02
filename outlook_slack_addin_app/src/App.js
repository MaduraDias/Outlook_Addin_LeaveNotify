import React, { Component } from 'react';
import './App.css';

import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import {Fabric} from 'office-ui-fabric-react/lib/Fabric';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';

class App extends Component {
  constructor(props) {
    super(props);
    
  }

  setSubject = (subject)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    //this.setState(subject)
    window.Office.context.mailbox.item.subject.setAsync(subject);
  };

   setToday = ()=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    window.Office.context.mailbox.item.start.setAsync( new Date());
     window.Office.context.mailbox.item.end.setAsync( new Date());
  };


   setStartdDate = (date)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    // this.setState({ value: date });
    window.Office.context.mailbox.item.start.setAsync( new Date(date));
  };

  setEndDate = (date)=>{
     // this.setState({ value: date });
    //window.Office.context.mailbox.item.subject = 'set the value';
    window.Office.context.mailbox.item.end.setAsync( new Date(date));
  };
  
  render() {
    return (
      <Fabric>
           
           <TextField  label='Subject' onChanged={this.setSubject} /> 
            <DatePicker label='Start Date' onSelectDate={this.setStartdDate} />
            <DatePicker label='End Date' onSelectDate={this.setEndDate} />
            <DefaultButton primary={ true } onClick={this.setToday} >
              Set Today
            </DefaultButton>
      </Fabric>
    );
  }
}

export default App;