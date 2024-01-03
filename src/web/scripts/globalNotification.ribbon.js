(function(scriptExport) {
    const notificationIds = {};

    function displayNotifications() {
        const userId = Xrm.Utility.getGlobalContext().userSettings.userId.replace("{", "").replace("}", "");

        return Xrm.WebApi.retrieveMultipleRecords("oss_notification", "?$select=oss_globalnotificationactionpayload&$filter=_ownerid_value eq " + userId + " and oss_globalnotificationactionpayload ne null")
        .then(function(results) {
            return results.entities
                .map(function(e) { return e.oss_globalnotificationactionpayload; })
                .filter(function(e) { return !!e; })
                .filter(function(value, index, self) {
                    return self.indexOf(value) === index;
                });
        })
        .then(function(messages) {
            messages.forEach(function(m) {
                const actionPayload = JSON.parse(m);
                
                if (notificationIds[actionPayload.message]) {
                    return;
                }

                const actionHandler = {
                    actionLabel: "View notifications",
                    eventHandler: function() {
                        const properties = {
                            pageType: "entitylist",
                            entityName: actionPayload.entityName
                        };

                        Xrm.Navigation.navigateTo(properties, { target: 1 })
                        .then(function() {
                            Xrm.App.clearGlobalNotification(notificationIds[actionPayload.message]);
                            delete notificationIds[actionPayload.message];
                        });
                    }
                };

                const notification = {
                    type: 2,
                    level: 4,
                    message: actionPayload.message,
                    showCloseButton: true,
                    action: actionHandler
                };

                Xrm.App.addGlobalNotification(notification)
                .then(function(result) {
                    notificationIds[actionPayload.message] = result;
                });
            });
        })
        .catch(console.error);
    };

    scriptExport.EnableRule = function() {
        try {
            displayNotifications()
                .catch(console.error);
        }
        catch(ex) {
            console.error(ex);
        }

        return false;
    }
})(window.Xoss_ApplicationRibbon_GlobalNotifications = window.Xoss_ApplicationRibbon_GlobalNotifications || {});