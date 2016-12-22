/// <reference path="Scripts/jquery-3.1.1.js" />
(function () {
    'use strict';

    Office.initialize = function (reason) {
        $(document).ready(function (reason) {
            /// Upon load connect to the server and request
            /// the sprints so that we can fill the select
            /// field on the form.
            makeAjaxCall("GetSprints", null, function (response) {
                /** @type {String} */
                var data = response.Message.toString();
                /** @type {String[]} */
                var results = data.split(",");
                $.each(results, function (index, value) {
                    $("#selectSprintList").append("<option>" + value + "</option>");
                })
            }, function (error) {
                $(document).append("<br/><p>" + error + "<p>");
            });

            makeAjaxCall("GetConfigs", null, function (response) {
                /** @type {String} */
                var data = response.Message.toString();
                /** @type {String[]} */
                var results = data.split(",");
                $.each(results, function (index, value) {
                    $("#selectConfigList").append("<option>" + value + "</option>");
                })
            });

            $("#launchSprintButton").click(function (event) {
                /// The user clicked the submit button. Build the URL
                /// from the selection in the select control and
                /// change the location
                var sprint = $("#selectSprintList").val();
                var config = $("#selectConfigList").val();
                var url = "";
                if (sprint != null && sprint != "" &&
                    config != null && sprint != "") {
                    // replace the launcher page with the proper sprint folder location,
                    // but be sure to keep all the query string values - to pass on
                    // to our OfficeJS page. However, add our one config setting to the end
                    url = window.location.href.replace("launcher.html", sprint + "/ComposeMessage.html");
                    url += "&Config=" + config;
                    // replace the url
                    location.replace(url);
                }
            });

        });
    }
})();

// Helper function to call the web service controller
makeAjaxCall = function (command, params, callback, error) {
    var dataToPassToService = {
        Command: command,
        Params: params
    };
    $.ajax({
        url: 'api/Default',
        type: 'POST',
        data: JSON.stringify(dataToPassToService),
        contentType: 'application/json;charset=utf-8',
        headers: { 'Access-Control-Allow-Origin': '*' },
        crossDomain: true
    }).done(function (data) {
        callback(data);
    }).fail(function (status) {
        error(status.statusText);
    })
};