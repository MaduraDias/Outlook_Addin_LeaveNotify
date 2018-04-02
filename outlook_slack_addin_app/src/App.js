import React, { Component } from 'react';
import './App.css';


class App extends Component {
  constructor(props) {
    super(props);
    
  }

  setSubject = (event)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    window.Office.context.mailbox.item.subject.setAsync(event.target.value);
  };

   setDate = (event)=>{
    //window.Office.context.mailbox.item.subject = 'set the value';
    window.Office.context.mailbox.item.start.setAsync( new Date("September 27, 2012 12:30:00"));
  };
  
  render() {
    return (
      <div id="content">
        <div id="content-header">
          <div className="padding">
            
            <label>Subject</label>
            <input type="text" name="subject" onChange={this.setSubject}/> 
            <button onClick={this.setDate} > Set Date </button>
          </div>
        </div>
      </div>
    );
  }
}

export default App;