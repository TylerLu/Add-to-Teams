(function () {
    // Create config and get AuthenticationContext
    window.config = {
        tenant: constant.tenant,
        clientId: constant.clientId,
        postLogoutRedirectUri: constant.postLogoutRedirectUri,
        cacheLocation: "localStorage",
        accessToken: "",
        isTeacher: false,
        url: window.location.href,
        userDisplayName: ""
    };
    var authContext = new AuthenticationContext(config);
    var isCallback = authContext.isCallback(window.location.hash);
    authContext.handleWindowCallback();

    window.onload = function () {
        getAccessToken();
    }

    /*
    Acquiring an Access Token
    */
    function getAccessToken() {
        var user = authContext.getCachedUser();
        if (!user) {
            authContext.login();
        }
        authContext.acquireToken(constant.graphApiUri, function (error, token) {
            if (error || !token) {
                console.log("ADAL error occurred: " + error);
                return;
            }
            else {
                config.accessToken = token;
                fetchUser();
                fetchTeams();
            }
        });
    }
    /*
    Get current user, user photo, check user is teacher or not  
    */
    function fetchUser() {
        // Fetch user's metadata.
        $.ajax({
            type: "GET",
            url: constant.graphApiUri + "/" + constant.graphVersion + "/me?$select=userPrincipalName,displayName",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            }
        }).done(function (data) {
            config.userDisplayName = (data.displayName == null ? data.userPrincipalName : data.displayName);
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
        xhr.open("GET", constant.graphApiUri + "/" + constant.graphVersion + "/me/photos/48x48/$value");
        xhr.setRequestHeader("Authorization", "Bearer " + config.accessToken);
        xhr.responseType = "blob";
        xhr.send();

        //Fetch user's roles to check if a teacher.
        $.ajax({
            type: "GET",
            url: constant.graphApiUri + "/" + constant.graphVersion + "/me/assignedLicenses",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            }
        }).done(function (data) {
            if (data && data.value) {
                $.each(data.value, function (i, val) {
                    if (val.skuId == constant.faculty) {
                        config.isTeacher = true;
                        return false;
                    }
                });
            }
        }).fail(function (error) {
            displayError(error);
        });
    }
    /*
    Fetch user's joined teams.
    */
    function fetchTeams() {
        $.ajax({
            type: "GET",
            url: constant.graphApiUri + "/" + constant.graphVersion + "/me/joinedTeams?$select=id,displayName",
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
        }).fail(function (error) {
            displayError(error);
        });
    }
    /*
    Post a assignment
    */
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
            url: constant.graphApiUri + "/" + constant.eduApiVersion + "/education/classes/" + config.teamId + "/assignments",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            },
            data: JSON.stringify(assignment),
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        }).done(function (data) {
            $("#panel2").addClass("d-none");
            addAssignmentResource(data);
        }).fail(function (error) {
            displayError(error);
        });
    }
    /*
    Add assignment resource
    */
    function addAssignmentResource(assignment) {
        // TODO: Use site's actual title as display name.
        var resource = {
            "resource": {
                "displayName": $(document).attr("title"),
                "link": config.url,
                "@odata.type": "#microsoft.education.assignments.api.educationLinkResource"
            }
        }
        $.ajax({
            type: "POST",
            url: constant.graphApiUri + "/" + constant.eduApiVersion + "/education/classes/" + config.teamId + "/assignments/" + assignment.id + "/resources",
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
    /*
    Publish assignment
    */
    function publishAssignment(assignment) {
        $.ajax({
            type: "POST",
            url: constant.graphApiUri + "/" + constant.eduApiVersion + "/education/classes/" + config.teamId + "/assignments/" + assignment.id + "/publish",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            },
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        }).done(function (data) {
            $("#panel3").removeClass("d-none");
            getAssignments();
        }).fail(function (error) {
            displayError(error);
        });
    }
    /*
    Get Assignments
    */
    function getAssignments() {
        $.ajax({
            type: "GET",
            url: constant.graphApiUri + "/" + constant.eduApiVersion + "/education/classes/" + config.teamId + "/assignments",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + config.accessToken
            }
        }).done(function (data) {
            if (data && data.value) {
                var html = "";
                $.each(data.value, function (i, val) {
                    html += "<tr>";
                    html += "<td>" + val.displayName + "</td>";
                    html += "<td>" + val.status + "</td>";
                    html += "<td>" + new Date(val.dueDateTime).toLocaleDateString() + "</td>";
                    html += "</tr>";
                });
                $("#assignmentsList").removeClass("d-none");
                $("#assignmentsList table tbody").html(html);
            }
        }).fail(function (error) {
            displayError(error);
        });
    }
    /*
    Post annoucement
    */
    function postAnnouncement(announcementText) {
        // TODO: Teams has a bug converting URLs to thumbnails.
        var announcement = {
            "rootMessage": {
                "body": {
                    "contentType": 1,
                    "content": "<div>" + announcementText + "</div>"
                }
            }
        }
        $.ajax({
            type: "POST",
            url: constant.graphApiUri + "/" + constant.graphVersion + "/groups/" + config.teamId + "/channels/" + config.channelId + "/chatthreads",
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
        }).fail(function (error) {
            displayError(error);
        });
    }
    /*
    Class Select Change
    */
    function onClassSelect() {
        var teamId = $("#selectClass").val();
        if (teamId) {
            config.teamId = teamId;
            showComponentByGroupChange("class");


            // Check if team is a class. Only a teacher can add assignment to a class.
            if (config.isTeacher) {
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
            }
        }
    }
    /*
    Action Select Change
    */
    function onActionSelect() {
        var actionId = $("#selectAction").val();
        if (actionId) {
            config.actionId = actionId;
            $.ajax({
                type: "GET",
                url: constant.graphApiUri + "/" + constant.graphVersion + "/groups/" + config.teamId + "/channels",
                headers: {
                    "Accept": "application/json",
                    "Authorization": "Bearer " + config.accessToken
                }
            }).done(function (data) {
                showComponentByGroupChange("action");
                data.value.forEach(function (channel) {
                    $("#selectChannel").append("<option value='" + channel.id + "'>" + channel.displayName + "</option>");
                });
            }).fail(function (error) {
                displayError(error);
            });
        }
    }
    /*
    Channel Select Change
    */
    function onChannelSelect() {
        var channelId = $("#selectChannel").val();
        if (channelId) {
            config.channelId = channelId;
            showComponentByGroupChange("channel");
        }
    }
    /*
    Show/Hide Dropdown when select change.
    */
    function showComponentByGroupChange(group) {
        if (group == "class") {
            $("#selectAction").empty();
            $("#selectAction").append("<option disabled selected>Choose an action</option>");
            $("#selectAction").append("<option value='announcement'>Make an announcement</option>");
            $("#selectAction").removeClass("d-none");

            $("#selectChannel").empty();
            $("#selectChannel").addClass("d-none");
            $("#button1").addClass("d-none");
        }
        else if (group == "action") {
            var actionId = $("#selectAction").val();

            $("#selectChannel").empty();
            $("#selectChannel").append("<option disabled selected>Choose a channel</option>");

            if (actionId && actionId == "announcement") {
                $("#selectChannel").removeClass("d-none");
                $("#button1").addClass("d-none");
            }
            else {
                $("#selectChannel").addClass("d-none");
                $("#button1").removeClass("d-none");
            }
        }
        else {
            $("#button1").removeClass("d-none");
        }
    }

    function onButton1Click() {
        // TODO: Replace with thumbnail and snippet generator.
        $("#thumbnail").prop("src", "/assets/khan-256.png");
        $("#caption").empty();
        $("#caption").append("<h5>" + config.userDisplayName + "</h5>");
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
    
    $("#selectClass").change(onClassSelect);
    $("#selectAction").change(onActionSelect);
    $("#selectChannel").change(onChannelSelect);
    $("#button1").click(onButton1Click);
    $("#button2").click(onButton2Click);
    $("#button3").click(onButton3Click);
})();
