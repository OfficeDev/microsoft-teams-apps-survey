The troubleshooting guide lists all known issues and workarounds for the issues.

## 1. Adaptive card does not get posted on hitting Create App

**Summary**: On creating a new app, the adaptive card for the app does not get posted on the chat canvas. 

**Fix**: Microsoft is trying to fix the issue. Until an update is released, make sure your tenants does not have a policy which convert multi-tenant AAD apps to single tenant AAD apps. More information on AAD apps can be found [here](https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-convert-app-to-be-multi-tenant#update-registration-to-be-multi-tenant%23:~:text=By%20default,Accounts%20in%20any%20organizational%20directory).

## 2. Deployment of app fails and throws an error of version doesn’t conform to ActionManifest.schema.json

**Summary**: On trying to deploy the app using the command `npm run deploy`, the below error shows  

ActionManifest doesn't conform to ActionManifest.schema.json. Errors: String '1.2' does not match regex pattern `^([0-9]|[1-9]+[0-9]*)\.([0-9]|[1-9]+[0-9]*)\.([0-9]|[1-9]+[0-9]*)$`. Path 'version' 

**Fix**: The schema for versioning follows Semantic version 2.0. The convention for version is MAJOR.MINOR.PATCH. Example of a version number is 1.1.2. 

## 3. Deployment of app fails and throws an error of URL format doesn’t conform to ActionManifest.schema.json

**Summary**: On trying to deploy the app using the command npm run deploy, the below error shows  

ActionManifest doesn't conform to ActionManifest.schema.json. Errors: String 'www.contoso.com' does not validate against format 'uri'. Path 'developer.privacyUrl'. 

**Fix**: Each URL must have a proper http or https scheme, like - https://www.contoso.com