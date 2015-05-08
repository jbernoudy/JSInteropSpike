/// <reference path="./typings/WinJS.d.ts"/>
/// <reference path="./typings/Salesforce.SDK.Hybrid.d.ts"/>
var Default;
(function (Default) {
    "use strict";
    var app = WinJS.Application;
    var activation = Windows.ApplicationModel.Activation;
    var oauth = new SalesforceJS.OAuth2();
    var auth = Salesforce.SDK.Hybrid.Auth;
    var oauth2 = Salesforce.SDK.Hybrid.Auth.OAuth2;
    var currentPage = "currentPage{0}";
    var defaultPage = "/one/one.app";
    var oneView;
    function navigateToOneApp() {
        var account = auth.HybridAccountManager.getAccount();
        var startPage = "";
        oauth2.refreshAuthToken(account).then(function (a) {
            account = a;
            startPage = getPage(account) + "?display=touch&sid=" + account.accessToken;
        });
        oauth2.refreshAuthToken(account).then(function (a) {
            account = a;
            oneView.navigate(startPage);
        });
    }
    function getPage(account) {
        var value = "";
        if (account !== null) {
            value = account.instanceUrl + defaultPage;
            var settings = Windows.Storage.ApplicationData.current.localSettings;
            var key = currentPage.replace("{0}", account.userId);
            if (settings.values.hasKey(key)) {
                value = settings.values[key];
            }
            if (value.search(defaultPage) === -1) {
                value = account.instanceUrl + defaultPage;
            }
        }
        return value;
    }
    function logout() {
        oauth.logout();
    }
    function startup() {
        oauth.configureOAuth("data/bootconfig.json", null)
            .then(function () {
            oauth.loginDefaultServer().done(function () {
                navigateToOneApp();
            }, function (error) {
                startup();
            });
        });
    }
    function savePage(account, page) {
        if (account !== null && page !== null && page.search(defaultPage) !== -1) {
            var settings = Windows.Storage.ApplicationData.current.localSettings;
            var key = currentPage.replace("{0}", account.userId);
            settings.values[key] = page;
        }
    }
    function scriptNotify(e) {
    }
    function navigationCompleted(e) {
        savePage(auth.HybridAccountManager.getAccount(), e.uri);
    }
    function logoutClicked() {
        logout();
    }
    app.addEventListener("activated", function (args) {
        logout();
        if (args.detail.kind === activation.ActivationKind.launch) {
            oneView = document.getElementById("oneView");
            oneView.addEventListener("MSWebViewScriptNotify", scriptNotify);
            oneView.addEventListener("MSWebViewFrameNavigationCompleted", navigationCompleted);
            if (args.detail.previousExecutionState !== activation.ApplicationExecutionState.terminated) {
                startup();
            }
            else {
                startup();
            }
            args.setPromise(WinJS.UI.processAll());
        }
    });
    app.oncheckpoint = function (args) {
        // TODO: This application is about to be suspended. Save any state
        // that needs to persist across suspensions here. You might use the
        // WinJS.Application.sessionState object, which is automatically
        // saved and restored across suspension. If you need to complete an
        // asynchronous operation before your application is suspended, call
        // args.setPromise().
    };
    app.start();
})(Default || (Default = {}));
//# sourceMappingURL=default.js.map