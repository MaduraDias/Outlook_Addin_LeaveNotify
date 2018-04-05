  
# **Office Outlook-Addin with Office UI Fabric React**

## **Technology Stack**
1) React - Office UI Fabric React - For the UI design
2) axios - [Promise based HTTP client for the browser and node.js](https://github.com/axios/axios) 



## **Prerequisites**

1)	Install create-react-app
```
npm install -g create-react-app
```

2)	Install Yeoman
```
npm install -g yo
```

3) Install Office Addin  Project Creator

```
npm install -g yo generator-office
```

## Creating the Addin App and Installing Dependencies (You have to follow the following steps only if you create a new add-in)

Run following commands on Command Prompt

1) Generate the new React App

```
 npx create-react-app outlook_leave_addin_app
```

2) Follow the steps in this [section](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react#generate-the-manifest-file-and-sideload-the-add-in) 

4)  Add styling
```
npm install --save @uifabric/styling
```
