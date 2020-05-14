function getProfileToken() {
    Office.context.auth.getAccessTokenAsync({allowConsentPrompt:true}, function(asyncResult) {
        if (asyncResult.status == "succeeded") {
            $("#profileToken").val(asyncResult.value);
            getGraphToken(asyncResult.value);
        } else {
            $("#profileToken").val(asyncResult.error.message);
        }

    });
}

function getGraphToken(profileToken) {
    $.ajax({type: "GET", 
		url: "/getpictures",
        headers: {"Authorization": "Bearer " + profileToken},
        cache: false
    }).then(function (response) {
        var list = $("<div class='list-group'>");
        for (var i = 0; i < response.length; i++) {
            var item = response[i];
            var listItem = $("<a href='#' class='list-group-item list-group-item-action'>").text(item.name);
            (function(downloadUrl) {
            listItem.click(function() {
                getImageContents(downloadUrl);
            });
            })(item.downloadUrl);
            list.append(listItem);
        }
        $("#fileResults").html(list);
    });
}

function getImageContents(downloadURL) {
    $.ajax({type: "GET", 
		url: "/getpictures/downloadpicture",
        headers: {"PictureUrl": downloadURL},
        cache: false
    }).then(function (response) {
        Excel.run(function(context) {
            var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
            shapes.addImage(response);
            return context.sync();
        });
    });
}

Office.initialize = function() {
    $(document).ready(function() {
        $("#getpictures").click(getProfileToken);
    });
}