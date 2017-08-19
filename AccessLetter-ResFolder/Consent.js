// Code in GitHub now

window.addEventListener("DOMContentLoaded", function () {
    SP.SOD.executeFunc("sp.js", "SP.ClientContext", init);
});
var itemID = null;
var cItem = null;
var regex = new RegExp("^[a-zA-Z0-9@_.+-]+$");
function init() {
    var encodedString = getParameterByName("identity");
    if (encodedString.length !== 0) {
        var decodedString = Base64.decode(encodedString);
        itemID = decodedString.split("=")[1];
        $("#btn_Submit").on("click", SubmitData);

    } else {
        console.log("Record does not exists");
    }
}

function getParameterByName(name) {
    name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
    var regex = new RegExp("[\\?&]" + name + "=([^&#]*)"),
        results = regex.exec(location.search);
    return results === null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
}

function SubmitData() {

    var vName = $("#txt_Name");
    var vEmail = $("#txt_Email");
    var vCompany = $("#txt_Company");

    if ((vName.val().length == 0) || (vEmail.val().length == 0) || (vCompany.val().length == 0)) {
        alert("Empty Field"); return;
        return false;
    }
    if (!validateEmail(vEmail.val())) {
        alert("Invalid Email"); return;
        return false;
    }
    $("#btn_Submit").attr("Value", "Accepting...");

    var context = new SP.ClientContext.get_current();
    var cItem = context.get_web().get_lists().getByTitle("Consent").getItemById(itemID);

    cItem.set_item("SName", $("#txt_Name").val());
    cItem.set_item("SEmail", $("#txt_Email").val());
    cItem.set_item("SCompany", $("#txt_Company").val());
    cItem.set_item("Consent", true);
    cItem.set_item("ConsentOn", new Date());
    cItem.set_item("ManagerInformed", false);
    cItem.update();
    context.executeQueryAsync(function () {
        $("#btn_Submit").attr("Value", "Accepted");
        $("#btn_Submit").attr("disabled", "true");

        $("#messageStatus").html("<img src='logo_overview.png' alt='bainlogo'><br><br>Thank you for accepting Bain & Company's terms of access. Your information is being processed and you should receive the report from the project manager shortly.<br><br> Please remember that the terms of access only allow you to share the report internally within your organization and its affiliates so if you want to pass the Materials to someone in another organization, they will also need to accept Bain & Company's terms of access before they can access the Materials.")
        $(".container-fluid").css("display", "none");
        var elmnt = document.getElementById("messageStatus");
        var y = elmnt.scrollTop;

    }, function (sender, args) {
        if (args.get_message() === "Item does not exist. It may have been deleted by another user.") {
            var camlQuery = new SP.CamlQuery();
            camlQuery.set_viewXml("<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + itemID + "</Value></Eq></Where></Query></View>");
            var VendorComebackRequestsItem = context.get_web().get_lists().getByTitle("VendorComebackRequests").getItems(camlQuery);
            context.load(VendorComebackRequestsItem, 'Include(Title)');
            context.executeQueryAsync(function () {
                if (VendorComebackRequestsItem.get_count() === 0) {
                    var itemCreateInfo = new SP.ListItemCreationInformation();
                    var oListItem = context.get_web().get_lists().getByTitle("VendorComebackRequests").addItem(itemCreateInfo);
                    oListItem.set_item("Title", itemID);
                    oListItem.set_item("SName", $("#txt_Name").val());
                    oListItem.set_item("SEmail", $("#txt_Email").val());
                    oListItem.set_item("SCompany", $("#txt_Company").val());
                    oListItem.set_item("ConsentOn", new Date());
                    oListItem.update();
                    context.executeQueryAsync(function () {
                        $(".container-fluid").css("display", "none");
                        $("#messageStatus").html("<img src='logo_overview.png' alt='bainlogo'><br><br>Thank you for accepting Bain & Company's terms of access. Your information is being processed and you should receive the report from the project manager shortly.<br><br> Please remember that the terms of access only allow you to share the report internally within your organization and its affiliates so if you want to pass the Materials to someone in another organization, they will also need to accept Bain & Company's terms of access before they can access the Materials.");
                        var elmnt = document.getElementById("messageStatus");
                        var y = elmnt.scrollTop;
                    }, function (sender, args) {
                        alert(args.get_message() + " ERROR in adding Vendor comeback request");
                        $("#btn_Submit").attr("Value", "Accepted");
                        $("#btn_Submit").attr("disabled", "true");
                    });
                } else if (VendorComebackRequestsItem.get_count() === 1) {
                    var itemToEdit = VendorComebackRequestsItem.get_item(0);
                    itemToEdit.set_item("SName", $("#txt_Name").val());
                    itemToEdit.set_item("SEmail", $("#txt_Email").val());
                    itemToEdit.set_item("SCompany", $("#txt_Company").val());
                    itemToEdit.set_item("ConsentOn", new Date());
                    itemToEdit.update();
                    context.executeQueryAsync(function () {
                        $(".container-fluid").css("display", "none");
                        $("#messageStatus").html("<img src='logo_overview.png' alt='bainlogo'><br><br>Thank you for accepting Bain & Company's terms of access. Your information is being processed and you should receive the report from the project manager shortly.<br><br> Please remember that the terms of access only allow you to share the report internally within your organization and its affiliates so if you want to pass the Materials to someone in another organization, they will also need to accept Bain & Company's terms of access before they can access the Materials.");
                        var elmnt = document.getElementById("messageStatus");
                        var y = elmnt.scrollTop;
                    }, function (sender, args) {
                        alert(args.get_message() + " ERROR in updating Vendor comeback request");
                        $("#btn_Submit").attr("Value", "Accepted");
                        $("#btn_Submit").attr("disabled", "true");
                    });
                }
            }, function (sender, args) {
                alert(args.get_message() + " ERROR connecting to VendorComebackRequests");
                $("#btn_Submit").attr("Value", "Accepted");
                $("#btn_Submit").attr("disabled", "true");
            });
        } else {
            $("#messageStatus").html(args.get_message());
            $("#btn_Submit").attr("Value", "Accepted");
            $("#btn_Submit").attr("disabled", "true");
        }
    });
}

function validateEmail($email) {
    var emailReg = /^([\w-\.]+@([\w-]+\.)+[\w-]{2,4})?$/;
    return emailReg.test($email);
}