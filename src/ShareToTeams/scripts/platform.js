﻿(function () {
    var elements = document.getElementsByClassName("sharetoteams");
    Array.prototype.forEach.call(elements, function (element) {
        var size = element.dataset.size;
        var url = element.dataset.url;

        var img = document.createElement("img");
        img.src = "https://sharetoteams.blob.core.windows.net/public/teams-" + size + ".png";
        img.onclick = function () {
            showPopUp("https://sharetoteams.azurewebsites.net/views/sharetoteams.html?url=" + encodeURIComponent(url), "_blank", 800, 600);
        }

        var a = document.createElement("a");
        a.href = "#";
        a.appendChild(img);

        element.appendChild(a);
    });

    function showPopUp(url, title, w, h) {
        var dualScreenLeft = window.screenLeft != undefined ? window.screenLeft : screen.left;
        var dualScreenTop = window.screenTop != undefined ? window.screenTop : screen.top;
        var width = window.innerWidth ? window.innerWidth : document.documentElement.clientWidth ? document.documentElement.clientWidth : screen.width;
        var height = window.innerHeight ? window.innerHeight : document.documentElement.clientHeight ? document.documentElement.clientHeight : screen.height;
        var left = (width / 2) - (w / 2) + dualScreenLeft;
        var top = (height / 2) - (h / 2) + dualScreenTop;
        var newWindow = window.open(url, title, "resizable=yes, width=" + w + ", height=" + h + ", top=" + top + ", left=" + left);
        if (window.focus) {
            newWindow.focus();
        }
    }
})();
