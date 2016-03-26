/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var _accessToken = {};
    var _dlg;

    Office.initialize = function (reason) {
        $(document).ready(function () {

            $('#btnGetPurchases').click(getPurchases);
            //$('#btnGetAccounts').click(getAccounts);
            $('#btnCreateReport').click(createDonationsTracker);
            $('#btnSignOut').click(signOut);
            $('#btnSignIn').click(signIn);
            $("#welcomePanel").show();
            $("#actionsPanel").hide();

            checkSignIn();
        });
    };

    function processMessage(arg) {
        _dlg.close();
        _accessToken = JSON.parse(arg.message);


        $.get("/api/setToken?t=" + _accessToken.token + "&s=" + _accessToken.secret)
            .done(function (data) {
                console.log(data);
            });

        $("#welcomePanel").hide();
        $("#actionsPanel").show();
    }

    function checkSignIn() {
        if (typeof _accessToken.token == "undefined") {
            $.get("api/getToken", function (data, status) {
                $("#welcomePanel").hide();
                $("#actionsPanel").fadeIn("slow");
            });
        }
    }


    function signIn() {
        Office.context.ui.displayDialogAsync("https://appcenter.intuit.com/Connect/SessionStart?grantUrl=https%3A%2F%2Flocalhost%3A44300%2FOAuthManager.aspx%3Fconnect%3Dtrue&datasources=quickbooks",
            { height: 40, width: 40, requireHTTPS: true },
            function (result) {
                _dlg = result.value;
                _dlg.addEventHandler("dialogMessageReceived", processMessage);
            });
    }


    function signOut() {
        $.get("/api/clearToken", function (data, status) {
            init();
        });
    }

    function getPurchases() {
        var url = "/api/getPurchases?n=100";

        $.get(url, function (data, status) {
            var tableBody = getFormattedArray(data);
            importTransactions(tableBody.data, tableBody.format);

        });

    }


    // Import sample transactions into the workbook
    function importTransactions(tableBodyData, tableBodyFormat) {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Unhide and start the spinner
            //$(".ms-Spinner").show();
            //spinnerComponent.start();

            // Add a new worksheet to store the transactions
            var dataSheet = ctx.workbook.worksheets.add("Transactions");

            // Create strings to store all the static content to display in the Welcome sheet
            var sheetTitle = "WoodGrove Bank";

            var sheetHeading1 = "Expense Transactions - Master List";

            var sheetDesc1 = "This is the master list of your spending activity.";

            var sheetDesc2 = "Filter transactions using the task pane to get insights.";

            var sheetDesc3 = "Track donations and flag items that need follow up.";

            var tableHeading = "Transactions";

            // Add all the intro content to the Welcome sheet and format the text
            addContentToWorksheet(dataSheet, "B1:E1", sheetTitle, "SheetTitle", "", "");

            addContentToWorksheet(dataSheet, "B3:E3", sheetHeading1, "SheetHeading", "", "");

            addContentToWorksheet(dataSheet, "C4:E4", sheetDesc1, "SheetHeadingDesc", "", "");

            addContentToWorksheet(dataSheet, "C5:E5", sheetDesc2, "SheetHeadingDesc", "", "");

            addContentToWorksheet(dataSheet, "C6:E6", sheetDesc3, "SheetHeadingDesc", "", "");

            addContentToWorksheet(dataSheet, "B19:B19", tableHeading, "TableHeading", "", "");

            //Fill white color in the sheet for improved look
            dataSheet.getRange("A2:I2000").format.fill.color = "white";

            // Queue a command to add a new table
            var startRowNumber = 20;
            var masterTableAddress = 'Transactions!B20:G' + (startRowNumber + tableBodyData.length).toString();
            var masterTable = ctx.workbook.tables.add(masterTableAddress, true);
            masterTable.name = "TransactionsTable";

            // Queue a command to get the newly added table
            masterTable.getHeaderRowRange().values = [["DATE", "AMOUNT", "MERCHANT", "CATEGORY", "TYPEOFDAY", "MONTH"]];

            masterTable.getDataBodyRange().formulas = tableBodyData;
            //masterTable.getDataBodyRange().numberFormat = tableBodyFormat;
            //masterTable.getDataBodyRange().format.autofitColumns();

            // Format the table header and data rows
            addContentToWorksheet(dataSheet, "B20:G20", "", "TableHeaderRow", "", "");

            addContentToWorksheet(dataSheet, "B21:G" + (startRowNumber + tableBodyData.length).toString(), "", "TableDataRows", "", "");

            // Set the number format of the Amount column
            masterTable.columns.getItem("AMOUNT").numberFormat = "$#,##0.00"; //'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';

            // Sort by most recent transactions at the top (Date, descending order)
            var sortRange = masterTable.getDataBodyRange().getColumn(0).getUsedRange();
            sortRange.sort.apply([
            {
                key: 0,
                ascending: false,
            },
            ]);

            //Queue a command to add the new chart
            var chartDataRangeColumn1 = masterTable.columns.getItemAt(0).getDataBodyRange();
            var chartDataRangeColumn2 = masterTable.columns.getItemAt(1).getDataBodyRange();
            // Comment about why we're doing this
            var chartDataRange = chartDataRangeColumn1.getBoundingRect(chartDataRangeColumn2);
            var chart = dataSheet.charts.add("Line", chartDataRange, Excel.ChartSeriesBy.auto);
            chart.setPosition("B8", "G17");
            chart.title.text = "Expense Trends";
            chart.title.format.font.color = "#41AEBD";
            chart.series.getItemAt(0).format.line.color = "#2E81AD";

            // Auto-fit columns and rows
            dataSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            dataSheet.getUsedRange().getEntireRow().format.autofitRows();


            // Set the sheet as active
            dataSheet.activate();

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        }).catch(function (error) {
            // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
            app.showNotification("Error: " + error);
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

    function getFormattedArray(purchases) {
        var result = {};
        result.data = [];
        result.format = [];

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

            result.data.push([date, amount, payee, cat, '=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")', '=TEXT([DATE], "mmm - yyyy")']);
            result.format.push([[null, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', null, null, null, null]]);
        });

        return result;
    }

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
			.catch(function (error) {
			    // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			    app.showNotification("Error: " + error);
			    console.log("Error: " + error);
			    if (error instanceof OfficeExtension.Error) {
			        console.log("Debug info: " + JSON.stringify(error.debugInfo));
			    }
			});
    }

    // Helper function to add and format content in the workbook
    function addContentToWorksheet(sheetObject, rangeAddress, displayText, typeOfText, formulaContent, numberFormat) {

        // Format differently by the type of content
        switch (typeOfText) {
            case "SheetTitle":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 30;
                range.format.font.color = "white";
                range.merge();
                //Fill color in the brand bar
                sheetObject.getRange("A1:M1").format.fill.color = "#41AEBD";
                break;
            case "SheetHeading":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 18;
                range.format.font.color = "#00b3b3";
                range.merge();
                break;
            case "SheetHeadingDesc":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.merge();
                break;
            case "SummaryDataHeader":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 13;
                range.merge();
                break;
            case "SummaryDataValue":
                var range = sheetObject.getRange(rangeAddress);
                range.numberFormat = numberFormat;
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 13;
                range.merge();
                break;
            case "TableHeading":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 12;
                range.format.font.color = "#00b3b3";
                range.merge();
                break;
            case "TableHeaderRow":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.format.font.bold = true;
                range.format.font.color = "black";
                break;
            case "TableDataRows":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                sheetObject.getRange(rangeAddress).format.borders.getItem('EdgeBottom').style = 'Continuous';
                sheetObject.getRange(rangeAddress).format.borders.getItem('EdgeTop').style = 'Continuous';
                break;
            case "TableTotalsRow":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.format.font.bold = true;
                break;
        }
    }

})();
