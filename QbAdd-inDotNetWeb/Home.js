/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    /*intuit.ipp.anywhere.setup({
        grantUrl: 'https://localhost:44300/api/qb?connect=true', datasources: {
            quickbooks: true,  // set to false if NOT using Quickbooks API
            payments: false    // set to true if using Payments API
        }
    });
    */
    var _accessToken = {};
    var _dlg;

    Office.initialize = function (reason) {
        $(document).ready(function () {
            //overrideWinOpen();

            $('#btnGetPurchases').click(getPurchases);
            //$('#btnGetAccounts').click(getAccounts);
            //$('#btnCreateReport').click(createReport);
            $('#btnSignOut').click(signOut);

            $('#btnSignIn').click(showDialog);

            init();
        });
    };

    function processMessage(arg) {
        _dlg.close();
        _accessToken = JSON.parse(arg.message);

        //$.post("/api/purchases/postToken", _accessToken).done(function (data) {
            //console.log(_accessToken);
        //});

            $.ajax({
                type: "GET",
                url: "/api/setToken?t=" + _accessToken.token + "&s=" + _accessToken.secret
            }).done(function (data) {
                console.log(data);
            });

        $("#welcomePanel").hide();
        $("#actionsPanel").show();
    }

    function init() {

        $.get("api/getToken", function (data, status) {
             if (data == "Success") {
                 $("#welcomePanel").hide();
                 $("#actionsPanel").fadeIn("slow");
             }
             else {
                 $("#welcomePanel").fadeIn("slow");
                 $("#actionsPanel").fadeIn("slow");
             }
         });
    }

    function signOut() {
        //Sign out from AD
        //var authContext = new AuthenticationContext(adalConfig);
        //authContext.clearCache();

        $.get("/clearToken", function (data, status) {
            init();
        });
    }


    function getPurchases() {
        var url = "/api/getPurchases?n=100"; //+ _accessToken.token + "&s=" + _accessToken.secret;

        $.get(url, function (data, status) {
            createPurchasesTable(data);
        });

    }

    function showDialog() {
        try {
            //Office.context.ui.displayDialogAsync("https://localhost:44300/initFlow.aspx",
            Office.context.ui.displayDialogAsync("https://appcenter.intuit.com/Connect/SessionStart?grantUrl=https%3A%2F%2Flocalhost%3A44300%2FOAuthManager.aspx%3Fconnect%3Dtrue&datasources=quickbooks",
                { height: 40, width: 40, requireHTTPS: true },
                function (result) {
                    _dlg = result.value;
                    _dlg.addEventHandler("dialogMessageReceived", processMessage);
                });
        } catch (e) {

            window.open("https://localhost:44300/test.aspx", "_bank")
        }

    }

    function createPurchasesTable(purchases) {
        Excel.run(function (ctx) {

            var sheet = ctx.workbook.worksheets.add("Expenses");
            sheet.activate();
            // Queue a command to add a new table
            var table = ctx.workbook.tables.add('Expenses!A2:E2', true);
            table.name = "Purchases";

            // Queue a command to get the newly added table
            table.getHeaderRowRange().values = [["Date", "Type", "Payee", "Category", "Amount"]];

            // Create a proxy object for the table rows
            var tableRows = table.rows;

            $.each(purchases, function (i, item) {
                var date = new Date(item.txnDateField).toLocaleDateString();
                var type = "Unspecified";
                switch (item.paymentTypeField) {
                    case 0:
                        type = "Cash";
                        break;
                    case 1:
                        type = "Check";
                        break;
                    case 2:
                        type = "Credit Card";
                        break;
                }
                var payee = "";
                if (item.entityRefField)
                    payee = item.entityRefField.nameField;
                var cat = "";
                if (item.lineField.length > 0) {
                    switch (item.lineField[0].detailTypeField) {
                        case 5:
                            cat = item.lineField[0].itemField.accountRefField.nameField;
                            break;
                        case "itemBasedExpenseLineDetailField":
                            cat = item.lineField[0].itemBasedExpenseLineDetailField.itemRefField.nameField;
                            break;

                    }
                }
                var amount = item.totalAmtField;

                var r = tableRows.add(null, [[date, type, payee, cat, amount]]);
                r.getRange().numberFormat = [[null, null, null, null, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)']];
                r.getRange().format.autofitColumns();

                addTitle(sheet, "A1:E1", "A1", "Expense");

            });



            return ctx.sync()

        }).catch(function (error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

    function addTitle(sheet, range, start, titleText) {

        var title = sheet.getRange(range);
        title.format.fill.color = "336699";
        title.format.font.color = "white";
        title.format.font.size = 24;
        title = sheet.getRange(start);
        title.values = titleText;

    }


})();
