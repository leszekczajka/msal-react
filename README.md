# README

Azure AD B2C is a identity provider covering social and enterprise logins.

The Microsoft Authentication library (MSAL) is JavaScript library to authenticate enterprise users using Microsoft Azure Active Directory (AAD), Microsoft account users (MSA), users using social identity providers like Facebook, Google, LinkedIn etc. and get access to Microsoft Cloud or Microsoft Graph.

This module is React wrapper over MSAL for Azure AD B2C. It fully supports active SSO sessions.

## Installation

If you are using npm:

    npm install @kdpw/msal-b2c-react --save

Or if you are using yarn:

    yarn add @kdpw/msal-b2c-react

## Initializing the Library

You'll first need to load the module and pass some configuration to the library. Normally this would go in your index.js file:

    import authentication from '@kdpw/msal-b2c-react';

    authentication.initialize({
        // optional, will default to this
        instance: 'https://login.microsoftonline.com/tfp/', 
        // your B2C tenant
        tenant: 'myb2ctenant.onmicrosoft.com',
        // the policy to use to sign in, can also be a sign up or sign in policy
        signInPolicy: 'mysigninpolicy',
        // the policy to use for password reset
        resetPolicy: 'mypasswordresetpolicy',
        // the the B2C application you want to authenticate with (that's just a random GUID - get yours from the portal)
        applicationId: '75ee2b43-ad2c-4366-9b8f-84b7d19d776e',
        // where MSAL will store state - localStorage or sessionStorage
        cacheLocation: 'sessionStorage',
        // the scopes you want included in the access token
        scopes: ['https://myb2ctenant.onmicrosoft.com/management/admin'],
        // optional, the redirect URI - if not specified MSAL will pick up the location from window.href
        redirectUri: 'http://localhost:3000',
        // optional, the URI to redirect to after logout
        postLogoutRedirectUri: 'http://myapp.com',
        // optional, default to true, set to false if you change instance
        validateAuthority: false
    });
    
## Authenticating When The App Starts

If you want to set things up so that a user is authenticated as soon as they hit your app (for example if you've got a link to an app from a landing page) then, in index.js, wrap the lines of code that launch the React app with the _authentication.run_ function:

    authentication.run(() => {
      ReactDOM.render(<App />, document.getElementById('root'));
      registerServiceWorker();  
    });

## Triggering Authentication Based on Components Mounting (and routing)

If you want to set things up so that a user is authenticated as they visit a part of the application that requires authentication then the appropriate components can be wrapped inside higher order components that will handle the authentication process. This is done using the _authentication.required_ function, normally in conjunction with a router. The example below shows this using the popular react-router:

    import React, { Component } from 'react';
    import authentication from '@kdpw/msal-b2c-react'
    import { BrowserRouter as Router, Route, Switch } from "react-router-dom";
    import HomePage from './Homepage'
    import MembersArea from './MembersArea'
    
    class App extends Component {
      render() {
        return (
          <Router basename={process.env.PUBLIC_URL}>
            <Switch>
              <Route exact path="/" component={HomePage} />
              <Route exact path="/membersArea" component={authentication.required(MembersArea)}>
            </Switch>
          </Router>
        );
      }
    }

## Getting the Id Token

Simply call the method _getIdToken_:

    import authentication from '@kdpw/msal-b2c-react'

    // ...

    const id_token = authentication.getIdToken();

## Getting the Access Token

Simply call the method _getAccessToken_:

    import authentication from '@kdpw/msal-b2c-react'

    // ...

    const access_token = authentication.getAccessToken();

## Getting the User Name

Simply call the method _getUserName_:

    import authentication from '@kdpw/msal-b2c-react'

    // ...

    const userName = authentication.getUserName();

## Signing Out

To sign out:

    import authentication from '@kdpw/msal-b2c-react'

    // ...

    authentication.signOut();

## Thanks

To build this I made fork of [react-azure-adb2c](https://github.com/JamesRandall/react-azure-adb2c) module. Thanks!
