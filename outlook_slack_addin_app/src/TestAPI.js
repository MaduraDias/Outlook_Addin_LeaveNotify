import React, { Component } from 'react';
import './App.css';


import  axios from 'axios';

class AppTestApi extends Component {


  
  constructor(props) {
    super(props);
     //this.setSubject("Leave Request");
    
    //this.getEmailList();
    
  }


  getEmailList = () => {

    // let setEmail = (response) => {
    //   this.setState({ emailList: response });
    // };

    var promise = axios.get('http://localhost:50525/api/Employee/6667/LeaveNotifyEmailList')
      .then(function (response) {
       console.log(response.data);
        //return response.data;

      }).catch(function (error) {
        console.log(error);
      });

  };

  render() {
    return (
      <div>
        <button onClick={this.getEmailList()}> Call Api </button>
      </div>
    );
  }
}

export default AppTestApi;