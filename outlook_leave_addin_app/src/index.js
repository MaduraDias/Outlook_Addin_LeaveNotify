import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';
import AppTestApi from './TestAPI';
import registerServiceWorker from './registerServiceWorker';

const Office = window.Office;

Office.initialize = () => {
    ReactDOM.render(<App />, document.getElementById('root'));
};
registerServiceWorker();
