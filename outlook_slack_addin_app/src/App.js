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

  setSubject = (value)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    window.Office.context.mailbox.item.subject.setAsync(value);
  };

   setDate = ()=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    window.Office.context.mailbox.item.start.setAsync( new Date("September 27, 2012 12:30:00"));
  };


   setSelectedDate = (value)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    window.Office.context.mailbox.item.start.setAsync( new Date(value));
  };
  
  render() {
    return (
      <Fabric>
           
           <TextField  label='Subject' onChanged={this.setSubject} /> 
            <DatePicker onSelectDate={this.setSelectedDate} />
            <DefaultButton primary={ true } onClick={this.setDate} >
              Change Date
            </DefaultButton>
      </Fabric>
    );
  }
}

export default App;