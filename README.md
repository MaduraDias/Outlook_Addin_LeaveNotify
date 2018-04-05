
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

## Creating the Addin App and Installing Dependencies 

Note : (You have to follow the following steps only if you create a new add-in

Run following commands on Command Prompt


1) Generate the new React App

```
 npx create-react-app outlook_leave_addin_app
```
2) Navigate to the outlook_leave_addin_app folder

```
cd outlook_leave_addin_app
```

3) Follow the steps [here](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react#generate-the-manifest-file-and-sideload-the-add-in) to create the office add-in manifest file.

4) Install "office-js".

```

npm install @micrsoft\office-js --save   
```
 Ref:(https://github.com/OfficeDev/office-js)

5) Install the Fabric React package 

```
npm --save install office-ui-fabric-react
```

6) For styling run following command to install required packages

```
npm install --save @uifabric/styling
```

7) Run the following command to use the [Office fabric UI Layouts](https://developer.microsoft.com/en-us/fabric#/styles/layout)

```
npm install office-ui-fabric-core
```
8) Run following command if you want to use axios.

```
npm install axios
```

Note: To run the React app locally with HTTPS type the following command in the location of the root folder.Beacuse, add-ins communicate only with https end points. 

```
set HTTPS=true&&npm start
```

References

https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/add-in-development-lifecycle

https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox.item

https://github.com/OfficeDev/office-js-docs/blob/master/reference/

https://github.com/OfficeDev/office-ui-fabric-react/tree/master/packages/office-ui-fabric-react/src/components

https://docs.microsoft.com/en-us/office/dev/add-ins/develop/use-the-oauth-authorization-framework-in-an-office-add-in

https://developer.microsoft.com/en-us/office/docs

https://github.com/OfficeDev/office-ui-fabric-react/blob/master/packages/icons/README.md

https://github.com/OfficeDev/office-ui-fabric-react/tree/master/packages/styling

https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react

https://github.com/OfficeDev/office-js

https://developer.microsoft.com/en-us/fabric#/get-started

