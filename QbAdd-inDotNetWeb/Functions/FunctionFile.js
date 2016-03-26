// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };

    function addDonation() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // First switch to the Transactions sheet in case it's not already active
            viewTransactionsSheet();

            // The new row to be added
            var rowToAdd;

            // Create a proxy object for the selected range and load its address and values properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, address");

            // Get the table
            var transactionsTable = ctx.workbook.tables.getItem("TransactionsTable");

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    rowToAdd = sourceRange.getEntireRow().getIntersection(transactionsTable.getRange());
                    rowToAdd.load("values");
                    rowToAdd.format.fill.color = "#92E4F0";
                })
                // Then run the queued-up commands, and return a promise to indicate task completion
                .then(ctx.sync)
                .then(function () {
                    // Get the donations sheet
                    var donationsSheet = ctx.workbook.worksheets.getItem("Donations");

                    // Get the donations table
                    var donationsTable = ctx.workbook.tables.getItem("DonationsTable");

                    // Create a proxy object for the table rows
                    var tableRows = donationsTable.rows;

                    // Queue commands to add some sample rows to the donations table
                    tableRows.add(null, [[rowToAdd.values[0][0], rowToAdd.values[0][1], rowToAdd.values[0][2], '=(TEXT([@DATE], "mmm - yyyy"))']]);

                    // Auto-fit columns and rows
                    donationsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
                    donationsSheet.getUsedRange().getEntireRow().format.autofitRows();

                    // Set the sheet as active
                    donationsSheet.activate();
                })
                // run ctx.sync here
                // Run the queued-up commands
            .then(ctx.sync)
            .then(updateDonationTrackerSummaryTables)
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


    function viewTransactionsSheet() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Get the transactions sheet
            var transactionsSheet = ctx.workbook.worksheets.getItem("Transactions");

            // Make the transactions sheet active
            transactionsSheet.activate();

            //Run the queued-up commands, and return a promise to indicate task completion
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
})();