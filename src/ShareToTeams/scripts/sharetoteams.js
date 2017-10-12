(function () {
    var config = {
        clientId: "20db89ee-263a-40d6-9256-103029570676",
        redirectUri: "http://addtoteamdev.azurewebsites.net/views/sharetoteams.html",
        scopes: ["User.Read", "User.Read.All", "Group.ReadWrite.All", "EduRoster.ReadBasic", "EduAssignments.ReadWriteBasic"],
        url: new URI().query(true).url
    }

    // Workaround to get MSAL to work with window.open. Figure this out with Azure.
    window.opener = null;

    // Fix for MSAL bug with Edge/IE.
    new Msal.Storage("localStorage");

    var userAgentApplication = new Msal.UserAgentApplication(config.clientId, null, requestTokenReceived, { redirectUri: config.redirectUri });

    window.onload = function () {
        if (!userAgentApplication.isCallback(window.location.hash) && window.parent === window) {
            getAccessToken();
        }
    }

    function requestTokenReceived(errorDesc, requestToken, error, tokenType) {
        if (errorDesc) {
            alert(errorDesc);
        } else {
            getAccessToken();
        }
    }

    function getAccessToken() {
        var user = userAgentApplication.getUser();
        if (!user) {
            // If user is not signed in, then prompt user to sign in via loginRedirect.
            userAgentApplication.loginRedirect(config.scopes);
        } else {
            // Try to acquire the token silently first.
            userAgentApplication.acquireTokenSilent(config.scopes)
                .then(function (accessToken) {
                    config.accessToken = accessToken;
                    fetchUser();
                    fetchTeams();
                }, function (error) {
                    // If acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
                    // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user 
                    // can re-type the current username and password and/ or give consent to new permissions your application is requesting.
                    // After authentication/authorization completes, this page will be reloaded again and getAccessToken() will be called.
                    // Then, acquireTokenSilent will then acquire the token silently and the Graph API call results will be made.
                    userAgentApplication.acquireTokenRedirect(config.scopes);
                });
        }
    }

    function fetchUser() {
        // Fetch user's metadata.
        $.ajax({
            type: "GET",
            url: "https://graph.microsoft.com/beta/me?$select=userPrincipalName",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            }
        }).done(function (data) {
            $("#username").html(data.userPrincipalName);
        }).fail(function (error) {
            displayError(error);
        });

        // Fetch user's picture. jQuery does not support fetching blobs.
        var xhr = new XMLHttpRequest();
        xhr.onreadystatechange = function () {
            if (this.readyState === 4 && this.status === 200) {
                var url = window.URL || window.webkitURL;
                $("#userphoto").attr("src", url.createObjectURL(this.response));
                $("#navbar").removeClass("d-none");
                $("#content").removeClass("d-none");
            }
        }
        xhr.open("GET", "https://graph.microsoft.com/beta/me/photos/48x48/$value");
        xhr.setRequestHeader("Authorization", "Bearer " + config.accessToken);
        xhr.responseType = "blob";
        xhr.send();
    }

    function fetchTeams() {
        // Fetch user's joined teams.
        $.ajax({
            type: "GET",
            url: "https://graph.microsoft.com/beta/me/joinedTeams?$select=id,displayName",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            }
        }).done(function (data) {
            $("#selectClass").empty();
            $("#selectClass").append("<option disabled selected>Choose a class</option>");

            data.value.forEach(function (team) {
                $("#selectClass").append("<option value='" + team.id + "'>" + team.displayName + "</option>");
            });

            $("#selectClass").change(onClassSelect);
        }).fail(function (error) {
            displayError(error);
        });
    }

    function postAnnouncement(announcementText) {
        // TODO: Teams has a bug converting URLs to thumbnails.
        var announcement = {
            "rootMessage": {
                "body": {
                    "contentType": 1,
                    "content": "<div><div>" + announcementText + "</div><div>" + config.url + "</div></div>"
                }
            }
        }

        $.ajax({
            type: "POST",
            url: "https://graph.microsoft.com/beta/groups/" + config.teamId + "/channels/" + config.channelId + "/chatthreads",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            },
            data: JSON.stringify(announcement),
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        }).done(function (data) {
            $("#panel2").addClass("d-none");
            $("#panel3").removeClass("d-none");
            $("#button3").click(onButton3Click);
        }).fail(function (error) {
            displayError(error);
        });
    }

    function postAssignment(assignmentName, assignmentDueDate) {
        // TODO: Allow adding of instructions. Currently a bug with Assignment's API.
        var assignment = {
            "displayName": assignmentName,
            "dueDateTime": assignmentDueDate,
            "status": "draft",
            "allowStudentsToAddResourcesToSubmission": true,
            "grading": {
                "@odata.type": "#microsoft.education.assignments.api.educationAssignmentPointsGradeType",
                "maxPoints": 100
            },
            "assignTo": {
                "@odata.type": "#microsoft.education.assignments.api.educationAssignmentClassRecipient"
            }
        }

        $.ajax({
            type: "POST",
            url: "https://graph.microsoft.com/testeduapi/education/classes/" + config.teamId + "/assignments",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            },
            data: JSON.stringify(assignment),
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        }).done(function (data) {
            $("#panel2").addClass("d-none");
            $("#panel3").removeClass("d-none");
            $("#button3").click(onButton3Click);

            addAssignmentResource(data);
        }).fail(function (error) {
            displayError(error);
        });
    }

    function addAssignmentResource(assignment) {
        // TODO: Use site's actual title as display name.
        var resource = {
            "resource": {
                "displayName": config.url,
                "link": config.url,
                "@odata.type": "#microsoft.education.assignments.api.educationLinkResource"
            }
        }

        $.ajax({
            type: "POST",
            url: "https://graph.microsoft.com/testeduapi/education/classes/" + config.teamId + "/assignments/" + assignment.id + "/resources",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            },
            data: JSON.stringify(resource),
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        }).done(function (data) {
            publishAssignment(assignment);
        }).fail(function (error) {
            displayError(error);
        });
    }

    function publishAssignment(assignment) {
        $.ajax({
            type: "POST",
            url: "https://graph.microsoft.com/testeduapi/education/classes/" + config.teamId + "/assignments/" + assignment.id + "/publish",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            },
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        }).done(function (data) {
        }).fail(function (error) {
            displayError(error);
        });
    }

    function onClassSelect() {
        var teamId = $("#selectClass").val();
        if (teamId) {
            config.teamId = teamId;
            $("#selectAction").empty();
            $("#selectAction").append("<option disabled selected>Choose an action</option>");
            $("#selectAction").append("<option value='announcement'>Make an announcement</option>");
            $("#selectAction").removeClass("d-none");
            $("#selectAction").change(onActionSelect);

            // Check if team is a class.
            $.ajax({
                type: "GET",
                url: "https://graph.microsoft.com/beta/groups/" + config.teamId + "?$select=extension_fe2174665583431c953114ff7268b7b3_Education_ObjectType",
                headers: {
                    "Accept": "application/json",
                    "Authorization": "Bearer " + config.accessToken
                }
            }).done(function (data) {
                if (data.extension_fe2174665583431c953114ff7268b7b3_Education_ObjectType === "Section") {
                    $("#selectAction").append("<option value='assignment'>Create an assignment</option>");
                }
            }).fail(function (error) {
                displayError(error);
            });

            // Fetch team's channels.
            $.ajax({
                type: "GET",
                url: "https://graph.microsoft.com/beta/groups/" + config.teamId + "/channels",
                headers: {
                    "Accept": "application/json",
                    "Authorization": "Bearer " + config.accessToken
                }
            }).done(function (data) {
                // TODO: Allow teachers to specify which channel to post to.
                var generalChannel = data.value.find(function (element) {
                    return element.displayName === "General";
                });

                if (generalChannel) {
                    config.channelId = generalChannel.id;
                } else {
                    alert("Could not find the General channel");
                }
            }).fail(function (error) {
                displayError(error);
            });
        }
    }

    function onActionSelect() {
        var actionId = $("#selectAction").val();
        if (actionId) {
            config.actionId = actionId;
            $("#button1").removeClass("d-none");
            $("#button1").click(onButton1Click);
        }
    }

    function onButton1Click() {
        // TODO: Replace with thumbnail and snippet generator.
        $("#thumbnail").prop("src", "https://sharetoteams.blob.core.windows.net/public/khan-256.png");
        $("#caption").empty();
        $("#caption").append("<h5>Khan Academy</h5>");
        $("#caption").append("<p>You can learn anything. Expert-created content and resources for every subject and level. Always free.</p>");

        if (config.actionId === "announcement") {
            $("#announcementInputs").removeClass("d-none");
        } else if (config.actionId === "assignment") {
            $("#assignmentDueDate").datepicker({
                todayHighlight: true,
                autoclose: true
            });
            $("#assignmentInputs").removeClass("d-none");
        }

        $("#panel1").addClass("d-none");
        $("#panel2").removeClass("d-none");
        $("#button2").click(onButton2Click);
    }

    function onButton2Click() {
        $("#button2").prop("disabled", true);

        if (config.actionId === "announcement") {
            // TODO: sanitize input.
            var announcementText = $("#announcementText").val();
            postAnnouncement(announcementText);
        } else if (config.actionId === "assignment") {
            var assignmentName = $("#assignmentName").val();
            var assignmentDueDate = new Date($("#assignmentDueDate input").val()).toISOString();
            postAssignment(assignmentName, assignmentDueDate);
        }
    }

    function onButton3Click() {
        window.close();
    }

    function displayError(error) {
        if (error.responseJSON.error.message) {
            alert(error.responseJSON.error.message);
        } else {
            alert(JSON.stringify(error));
        }
    }
})();
