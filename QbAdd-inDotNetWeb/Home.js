/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var accessToken = {};
    var dlg;
    var messageBanner = null;

    Office.initialize = function (reason) {
        $(document).ready(function () {

            $('#btnSignIn').click(signIn);
            $('#btnSignOut').click(signOut);

            $('#btnGetExpenses').click(getExpenses);
            //$('#btnGetAccounts').click(getAccounts);
            $('#btnCreateReport').click(createDonationsTracker);

            $("#welcomePanel").show();
            $("#actionsPanel").hide();

            // Initialize the FabricUI notification mechanism and hide it
            messageBanner = new fabric.MessageBanner($(".ms-MessageBanner")[0]);
            messageBanner.hideBanner();

            checkSignIn();
        });
    };

    function signIn() {
        var signInUrl = "https://appcenter.intuit.com/Connect/SessionStart?datasources=quickbooks&grantUrl="
            //+ encodeURIComponent("https://localhost:44300/OAuthManager.aspx?connect=true");
            + encodeURIComponent("https://qbaddin.azurewebsites.net/OAuthManager.aspx?connect=true");
        Office.context.ui.displayDialogAsync(signInUrl,
            { height: 40, width: 40},
            function (result) {
                dlg = result.value;
                dlg.addEventHandler("dialogMessageReceived", processMessage);
            });
    }

    function processMessage(arg) {
        dlg.close();
        accessToken = JSON.parse(arg.message);

        $.get("/api/setToken?" + $.param(accessToken))
            .done(function (data) {
                console.log(data);
            });

        $("#welcomePanel").hide();
        $("#actionsPanel").show();
    }

    function signOut() {
        $.get("/api/clearToken", function (data, status) {
            accessToken = null;
            $("#welcomePanel").show();
            $("#actionsPanel").hide();
        });
    }

    function checkSignIn() {
        if (typeof accessToken.token == "undefined") {
            $.get("api/getToken", function (data, status) {
                $("#welcomePanel").hide();
                $("#actionsPanel").fadeIn("slow");
            });
        }
    }


    function getExpenses() {
        var url = "/api/getExpenses?n=100";

        $.get(url, function (data, status) {
            var tableBody = getFormattedArray(data);
            addExpensesSheet(tableBody);
        });
    }


    // Import sample transactions into the workbook
    function addExpensesSheet(tableBody) {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Add a new worksheet to store the transactions
            var dataSheet = ctx.workbook.worksheets.add("Expenses");

           //Fill white color in the sheet for improved look
            dataSheet.getRange("A2:M1000").format.fill.color = "white";

            //Add Sheet Title
            var range = dataSheet.getRange("B1:E1");
            range.values = "Contoso Expenses";
            range.format.font.name = "Corbel";
            range.format.font.size = 30;
            range.format.font.color = "white";
            range.merge();
            //Fill color in the brand bar
            dataSheet.getRange("A1:M1").format.fill.color = "#41AEBD";

            // Queue a command to add a new table
            var startRowNumber = 2;
            var masterTableAddress = 'Expenses!B' + startRowNumber + ':G' + (startRowNumber + tableBody.data.length).toString();
            var masterTable = ctx.workbook.tables.add(masterTableAddress, true);
            masterTable.name = "ExpensesTable";

            // Queue a command to get the newly added table
            masterTable.getHeaderRowRange().values = [["DATE", "AMOUNT", "MERCHANT", "CATEGORY", "TYPEOFDAY", "MONTH"]];

            masterTable.getDataBodyRange().formulas = tableBody.data;
            masterTable.getDataBodyRange().numberFormat = tableBody.format;

            // Format the table header and data rows
            range = dataSheet.getRange('Expenses!B' + startRowNumber + ':G' + startRowNumber);
            range.format.font.name = "Corbel";
            range.format.font.size = 10;
            range.format.font.bold = true;
            range.format.font.color = "black";

            range = dataSheet.getRange(masterTableAddress);
            range.format.font.name = "Corbel";
            range.format.font.size = 10;
            range.format.borders.getItem('EdgeBottom').style = 'Continuous';
            range.format.borders.getItem('EdgeTop').style = 'Continuous';

            // Sort by most recent transactions at the top (Date, descending order)
            var sortRange = masterTable.getDataBodyRange().getColumn(0).getUsedRange();
            sortRange.sort.apply([
            {
                key: 0,
                ascending: false,
            },
            ]);

            // Auto-fit columns and rows
            dataSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            dataSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Set the sheet as active
            dataSheet.activate();

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        }).catch(errorHandler);

    }

    function getFormattedArray(expenses) {
        var result = {};
        result.data = [];
        result.format = [];

        $.each(expenses, function (i, item) {
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

            result.data.push([date, amount, payee, cat, '=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")', '=TEXT([DATE], "mmm - yyyy")']);
            result.format.push([null, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', null, null, null, null]);
        });

        return result;
    }

    // Helper function to add and format content in the workbook
 
    // Create the charitable donations tracker
    function createDonationsTracker() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Add a new worksheet to store the transactions
            var donationsSheet = ctx.workbook.worksheets.add("Donations");

            // Create strings to store all the static content to display in the Welcome sheet
            var sheetTitle = "WoodGrove Bank";

            var sheetHeading1 = "Donations Tracker";

            var sheetDesc1 = "Track your charitable contributions throughout the year.";

            var sheetDesc2 = "Use this data at the end of the year to report your tax deductions.";

            var sheetHeading2 = "Summary";

            var summaryDataHeader1 = "Total Donations";

            var tableHeading1 = "Donations By Organization";

            var tableHeading2 = "Donations By Month";

            var tableHeading3 = "Transaction Details";

            // Add all the intro content to the Welcome sheet and format the text
            addContentToWorksheet(donationsSheet, "B1:G1", sheetTitle, "SheetTitle", "", "");

            addContentToWorksheet(donationsSheet, "B3:C3", sheetHeading1, "SheetHeading", "", "");

            addContentToWorksheet(donationsSheet, "C4:G4", sheetDesc1, "SheetHeadingDesc", "", "");

            addContentToWorksheet(donationsSheet, "C5:J5", sheetDesc2, "SheetHeadingDesc", "", "");

            addContentToWorksheet(donationsSheet, "B7:B7", sheetHeading2, "SheetHeading", "", "");

            addContentToWorksheet(donationsSheet, "B9:C9", summaryDataHeader1, "SummaryDataHeader", "", "");

            addContentToWorksheet(donationsSheet, "B11:D11", tableHeading1, "TableHeading", "", "");

            addContentToWorksheet(donationsSheet, "E11:F11", tableHeading2, "TableHeading", "", "");

            addContentToWorksheet(donationsSheet, "H11:K11", tableHeading3, "TableHeading", "", "");

            //Fill white color in the sheet for improved look
            donationsSheet.getRange("A2:L250").format.fill.color = "white";

            // Queue a command to add the Transaction Details table
            var donationsTable = ctx.workbook.tables.add('Donations!H12:K12', true);
            donationsTable.name = "DonationsTable";

            // Queue a command to get the newly added table
            donationsTable.getHeaderRowRange().values = [["DATE", "AMOUNT", "ORGANIZATION", "MONTH"]];

            // Queue a command to add the Summary Donations by Organization table
            var donationsByOrgTable = ctx.workbook.tables.add('Donations!B12:C12', true);
            donationsByOrgTable.name = "DonationsByOrgTable";

            // Queue a command to get the newly added table
            donationsByOrgTable.getHeaderRowRange().values = [["ORGANIZATION", "AMOUNT"]];
            donationsByOrgTable.showTotals = true;
            donationsByOrgTable.getTotalRowRange().getLastCell().values = [["=SUM([AMOUNT]"]];

            // Queue a command to add the Summary Donations by Month table
            var donationsByMonthTable = ctx.workbook.tables.add('Donations!E12:F12', true);
            donationsByMonthTable.name = "DonationsByMonthTable";

            // Queue a command to get the newly added table
            donationsByMonthTable.getHeaderRowRange().values = [["MONTH", "AMOUNT"]];
            donationsByMonthTable.showTotals = true;
            donationsByMonthTable.getTotalRowRange().getLastCell().values = [["=SUM([AMOUNT]"]];

            addContentToWorksheet(donationsSheet, "B12:C12", "", "TableHeaderRow", "", "");

            addContentToWorksheet(donationsSheet, "B13:C250", "", "TableDataRows", "", "");

            addContentToWorksheet(donationsSheet, "E12:F12", "", "TableHeaderRow", "", "");

            addContentToWorksheet(donationsSheet, "E13:F250", "", "TableDataRows", "", "");

            addContentToWorksheet(donationsSheet, "H12:K12", "", "TableHeaderRow", "", "");

            addContentToWorksheet(donationsSheet, "G13:J250", "", "TableDataRows", "", "");

            // Set the number format for Date and Currency columns
            donationsSheet.getRange("C13:C200").numberFormat = "$#";
            donationsSheet.getRange("F13:F200").numberFormat = "$#";
            donationsSheet.getRange("I13:I200").numberFormat = "$#";
            donationsSheet.getRange("E13:E200").numberFormat = "@";
            donationsSheet.getRange("K13:K200").numberFormat = "mmm-yyyy";
            donationsSheet.getRange("H13:H200").numberFormat = "mm/dd/yyyy";
            donationsByOrgTable.getTotalRowRange().getLastCell().numberFormat = "$#";
            donationsByMonthTable.getTotalRowRange().getLastCell().numberFormat = "$#";

            // Set the value of Total Donations at the top
            var rangetotalDonated = donationsSheet.getRange("D9");
            rangetotalDonated.formulas = [["=SUM(DonationsTable[AMOUNT])"]];
            rangetotalDonated.format.font.name = "Corbel";
            rangetotalDonated.format.font.size = 18;
            rangetotalDonated.numberFormat = "$#";
            rangetotalDonated.merge();

            // Auto-fit columns and rows
            donationsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            donationsSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
			.catch(errorHandler);
    }


    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
