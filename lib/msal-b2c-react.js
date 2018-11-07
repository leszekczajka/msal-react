'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _msal = require('msal');

var Msal = _interopRequireWildcard(_msal);

var _react = require('react');

var _react2 = _interopRequireDefault(_react);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

function _interopRequireWildcard(obj) { if (obj && obj.__esModule) { return obj; } else { var newObj = {}; if (obj != null) { for (var key in obj) { if (Object.prototype.hasOwnProperty.call(obj, key)) newObj[key] = obj[key]; } } newObj.default = obj; return newObj; } }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; } // note on window.msal usage. There is little point holding the object constructed by new Msal.UserAgentApplication
// as the constructor for this class will make callbacks to the acquireToken function and these occur before
// any local assignment can take place. Not nice but its how it works.


var logger = new Msal.Logger(loggerCallback, {
  level: Msal.LogLevel.Warning
});
var state = {
  noScopes: false,
  launchApp: null,
  idToken: null,
  accessToken: null,
  userName: ""
};
var appConfig = {
  // optional, will default to 'https://login.microsoftonline.com/tfp/'
  instance: null,
  // your B2C tenant
  tenant: null,
  // the policy to use to sign in, can also be a sign up or sign in policy
  signInPolicy: null,
  // the policy to use for password reset
  resetPolicy: null,
  // the the B2C application you want to authenticate with
  applicationId: null,
  // where MSAL will store state - localStorage or sessionStorage
  cacheLocation: null,
  // optional, the scopes you want included in the access token
  scopes: [],
  // optional, the redirect URI - if not specified MSAL will pick up the location from window.href
  redirectUri: null,
  // optional, the URI to redirect to after logout
  postLogoutRedirectUri: null,
  // optional, default to true, set to false if you change instance
  validateAuthority: null
};

function loggerCallback(logLevel, message, piiLoggingEnabled) {
  console.log(message);
}

function authCallback(errorDesc, token, error, tokenType) {
  if (errorDesc && errorDesc.indexOf('AADB2C90118') > -1) {
    redirect();
  } else if (errorDesc) {
    console.log(error + ':' + errorDesc);
  } else {}
}

function redirect() {
  var localMsalApp = window.msal;
  var instance = appConfig.instance ? appConfig.instance : 'https://login.microsoftonline.com/tfp/';
  var authority = '' + instance + appConfig.tenant + '/' + appConfig.resetPolicy;
  localMsalApp.authority = authority;
  loginAndAcquireToken();
}

function loginAndAcquireToken(successCallback) {
  var localMsalApp = window.msal;
  var user = localMsalApp.getUser(appConfig.scopes);

  if (!user) {

    // user is not logged in
    if (state.noScopes) {
      // no need of access token, just redirect to login page
      localMsalApp.loginRedirect(appConfig.scopes);
    } else {
      // try to get token from SSO session
      localMsalApp.acquireTokenSilent(appConfig.scopes, null, null, "&login_hint&domain_hint=organizations").then(function (accessToken) {
        state.accessToken = accessToken;
        user = localMsalApp.getUser(appConfig.scopes);
        state.idToken = user.idToken;
        state.userName = user.name;
        if (state.launchApp) {
          state.launchApp();
        }
        if (successCallback) {
          successCallback();
        }
      }, function (error) {
        if (error) {
          localMsalApp.loginRedirect(appConfig.scopes);
        }
      });
    }
  } else {

    // the user is already logged in
    state.idToken = user.idToken;
    state.userName = user.name;
    if (state.noScopes) {
      // no need of access token, just launch the app
      if (state.launchApp) {
        state.launchApp();
      }
      if (successCallback) {
        successCallback();
      }
    } else {
      // get access token
      localMsalApp.acquireTokenSilent(appConfig.scopes).then(function (accessToken) {
        state.accessToken = accessToken;
        if (state.launchApp) {
          state.launchApp();
        }
        if (successCallback) {
          successCallback();
        }
      }, function (error) {
        if (error) {
          localMsalApp.acquireTokenRedirect(appConfig.scopes);
        }
      });
    }
  }
}

var authentication = {
  initialize: function initialize(config) {
    appConfig = config;
    var instance = config.instance ? config.instance : 'https://login.microsoftonline.com/tfp/';
    var authority = '' + instance + config.tenant + '/' + config.signInPolicy;
    var validateAuthority = config.validateAuthority != null ? config.validateAuthority : true;
    var scopes = config.scopes;
    if (!scopes || scopes.length === 0) {
      console.log('To obtain access tokens you must specify one or more scopes. See https://docs.microsoft.com/en-us/azure/active-directory-b2c/active-directory-b2c-access-tokens');
      state.noScopes = true;
    }
    state.scopes = scopes;

    new Msal.UserAgentApplication(config.applicationId, authority, authCallback, {
      logger: logger,
      cacheLocation: config.cacheLocation,
      postLogoutRedirectUri: config.postLogoutRedirectUri,
      redirectUri: config.redirectUri,
      validateAuthority: validateAuthority
    });
  },
  run: function run(launchApp) {
    state.launchApp = launchApp;
    if (!window.msal.isCallback(window.location.hash) && window.parent === window && !window.opener) {
      loginAndAcquireToken();
    }
  },
  required: function required(WrappedComponent, renderLoading) {
    return function (_React$Component) {
      _inherits(_class, _React$Component);

      function _class(props) {
        _classCallCheck(this, _class);

        var _this = _possibleConstructorReturn(this, (_class.__proto__ || Object.getPrototypeOf(_class)).call(this, props));

        _this.state = {
          signedIn: false,
          error: null
        };
        return _this;
      }

      _createClass(_class, [{
        key: 'componentWillMount',
        value: function componentWillMount() {
          var _this2 = this;

          loginAndAcquireToken(function () {
            _this2.setState(Object.assign({}, _this2.state, {
              signedIn: true
            }));
          });
        }
      }, {
        key: 'render',
        value: function render() {
          if (this.state.signedIn) {
            return _react2.default.createElement(WrappedComponent, this.props);
          };
          return typeof renderLoading === 'function' ? renderLoading() : null;
        }
      }]);

      return _class;
    }(_react2.default.Component);
  },
  signOut: function signOut() {
    window.msal.logout();
  },
  getIdToken: function getIdToken() {
    return state.idToken;
  },
  getAccessToken: function getAccessToken() {
    return state.accessToken;
  },
  getUserName: function getUserName() {
    return state.userName;
  }
};

exports.default = authentication;