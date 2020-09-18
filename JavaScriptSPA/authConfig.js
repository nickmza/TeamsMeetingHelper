 
// Config object to be passed to Msal on creation.
// For a full list of msal.js configuration parameters, 
// visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
const msalConfig = {
  auth: {
    clientId: "d275b0c1-b498-4ab4-89b2-6e3aefbf6c23",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://meetinghelper.azurewebsites.net/",
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  }
};  
  
// Add here scopes for id token to be used at MS Identity Platform endpoints.
const loginRequest = {
  scopes: ["openid", "profile", "User.Read","OnlineMeetings.ReadWrite"]
};

// Add here scopes for access token to be used at MS Graph API endpoints.
const tokenRequest = {
  //scopes: ["Mail.Read"]
  scopes: ["OnlineMeetings.ReadWrite"]
};