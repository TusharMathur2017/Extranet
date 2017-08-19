var context = null
var currentUser = null;
var excelTab = null;
window.addEventListener("DOMContentLoaded", function () {
    SP.SOD.executeFunc("sp.js", "SP.ClientContext", init);
});

function init() {
    var signInLink = $(".ms-signInLink")[0];
    if (signInLink != undefined) {
        if (signInLink.innerHTML === "Sign In") {
            window.open(_spPageContextInfo.siteAbsoluteUrl + "/_layouts/15/Authenticate.aspx?Source=" + window.location.pathname, "_self", true);
        }
    }
    $("#txt_VCount").focusout(function (e) {
        if (parseInt($(this).val()) > 0) {
            excelTab = $("#excelControl")
                .empty()
                .xtab("init", {
                    rows: this.value,
                    cols: 3,
                    collabels: false,
                    rowlabels: false,
                    widths: [159, 159, 159],
                });
            $(".hideComp").css("visibility", "visible");
        } else {
            $(".hideComp").css("visibility", "hidden");
            alert("Must enter a number to continue");
        }
    });
    $("#txt_vMultiLine").focusout(function (e) {
        var rowCount = 0;
        var data = $('#txt_vMultiLine').val();
        var rows = data.split("\n");
        for (var y in rows) {
            var cells = rows[y].split("\t");
            for (var x in cells) {
                $("#excelControl-" + rowCount + "-" + x).val(cells[x]);
            }
            rowCount++;
        }
        $('#txt_vMultiLine').val("");
    });
    $("input[type=radio][name=Diligence]").change(function () {
        if (this.value === "Yes") {
            alert("Please be aware that you may only share VDD reports in the second round of a sales process and with a maximum of 6 bidders.\n\nIf you have any questions or concerns about this, please contact a member of the EMEA legal team");
        }
    });
    $("#btn_Send").on("click", function () {
        context = new SP.ClientContext.get_current();
        currentUser = context.get_web().get_currentUser();
        context.load(currentUser);
        context.executeQueryAsync(function () {
            SubmitData(excelTab.xtab("val"));
        }, function (sender, args) {
            alert("Error getting current user email.. " + args.get_message());
        });
    });
}

function SubmitData(dataExcelArea) {
    var newItemsArray = [];
    var rowCount = 0;
    var allRec = true;
    var invalidData = false;
    var senderList = "\n";

    if ($("#txt_CaseCode")[0].value.length === 0 ||
        $("#txt_ClientName")[0].value.length === 0 ||
        $("#txt_PName")[0].value.length === 0 ||
        $("#txt_VCount")[0].value.length === 0) {
        alert("Please fill in all the values before submitting the form");
        return;
    }

    for (var index = 0; index < dataExcelArea.length; index++) {
        var element = dataExcelArea[index];
        var newEntryItem = context.get_web().get_lists().getByTitle("Consent").addItem(new SP.ListItemCreationInformation());
        if ((element[0].length != 0) &&
            (element[1].length != 0) &&
            (element[2].length != 0)) {
            if (validateEmail(element[0])) {
                newEntryItem.set_item("Title", element[0]);
                newEntryItem.set_item("RName", element[1]);
                newEntryItem.set_item("RCompany", element[2]);
                newEntryItem.set_item("CaseCode", $("#txt_CaseCode")[0].value.trim());
                newEntryItem.set_item("ClientName", $("#txt_ClientName")[0].value.trim());
                newEntryItem.set_item("ProjectName", $("#txt_PName")[0].value.trim());
                if ($("input:radio[name=Diligence]:checked").val() === "Yes") {
                    newEntryItem.set_item("VendorDiligence", true);
                } else {
                    newEntryItem.set_item("VendorDiligence", false);
                }
                newEntryItem.update();
                newItemsArray[rowCount] = newEntryItem;
                context.load(newItemsArray[rowCount]);
                rowCount++;
                senderList += element[0] + "\n";
            }
            else {
                alert(element[0] + " is NOT a valid email address");
                return false;
            }
        } else {
            invalidData = true;
            var allRec = false;
            alert("Values missing from entry: " + element[0] + " " + element[1] + " " + element[2]);
        }
    };
    if (invalidData === false) {
        $("#btn_Send").attr("Value", "Sending...");
    }
    else {
        return;
    }
    context.executeQueryAsync(function () {
        for (var i = newItemsArray.length - 1; i >= 0; i--) {
            var oListItem = context.get_web().get_lists().getByTitle("Consent").getItemById(newItemsArray[i].get_id());
            var eURL = "?identity=" + newItemsArray[i].get_id();
            var encodedString = Base64.encode(eURL);
            oListItem.set_item("ConsentURL", encodedString);
            oListItem.update();
        }
        context.executeQueryAsync(function () { }, function (sender, args) {
            $("#btn_Send").attr("Value", "Send Email");
            $("#btn_Send").attr("disabled", "true");
            alert("Error creating encrypted URL's... " + args.get_message());
        });
    }, function (sender, args) {
        allRec = false;
        $("#btn_Send").attr("Value", "Send Email");
        $("#btn_Send").attr("disabled", "true");
        alert("Error storing data... " + args.get_message());
    });
    if (allRec) {
        $("#txt_CaseCode")[0].value = "";
        $("#txt_ClientName")[0].value = "";
        $("#txt_PName")[0].value = "";
        $("#txt_VCount")[0].value = "";
        $("#btn_Send").attr("Value", "Send Email");
        $(".hideComp").css("visibility", "hidden");
        $("#excelControl").empty();
        alert("Thank you - request e-mails have been sent to the below recipients asking them to accept Bain's terms. After each recipient has accepted via the system, you will get an e-mail informing you that you can send the report.\n" +
            senderList + "\n" +
            "Thank you for using the Bain Report Online Access Authorisation Centre");
    }
}

function validateEmail($email) {
    var emailReg = /^([\w-\.]+@([\w-]+\.)+[\w-]{2,4})?$/;
    return emailReg.test($email);
}