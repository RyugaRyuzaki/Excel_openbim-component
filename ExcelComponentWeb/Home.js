import { initialize } from "./Scripts/loadViewer.js";
(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded.
	// make sure office is initialized
	Office.initialize = function (reason) {};
	document.addEventListener("DOMContentLoaded", initialize());

	// now we initialize viewer

	// let test in browser first

	// function loadSampleData() {
	//     var values = [
	//         [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
	//         [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
	//         [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
	//     ];

	//     // Run a batch operation against the Excel object model
	//     Excel.run(function (ctx) {
	//         // Create a proxy object for the active sheet
	//         var sheet = ctx.workbook.worksheets.getActiveWorksheet();
	//         // Queue a command to write the sample data to the worksheet
	//         sheet.getRange("B3:D5").values = values;

	//         // Run the queued-up commands, and return a promise to indicate task completion
	//         return ctx.sync();
	//     })
	//     .catch(errorHandler);
	// }

	// function hightlightHighestValue() {
	//     // Run a batch operation against the Excel object model
	//     Excel.run(function (ctx) {
	//         // Create a proxy object for the selected range and load its properties
	//         var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

	//         // Run the queued-up command, and return a promise to indicate task completion
	//         return ctx.sync()
	//             .then(function () {
	//                 var highestRow = 0;
	//                 var highestCol = 0;
	//                 var highestValue = sourceRange.values[0][0];

	//                 // Find the cell to highlight
	//                 for (var i = 0; i < sourceRange.rowCount; i++) {
	//                     for (var j = 0; j < sourceRange.columnCount; j++) {
	//                         if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
	//                             highestRow = i;
	//                             highestCol = j;
	//                             highestValue = sourceRange.values[i][j];
	//                         }
	//                     }
	//                 }

	//                 cellToHighlight = sourceRange.getCell(highestRow, highestCol);
	//                 sourceRange.worksheet.getUsedRange().format.fill.clear();
	//                 sourceRange.worksheet.getUsedRange().format.font.bold = false;

	//                 // Highlight the cell
	//                 cellToHighlight.format.fill.color = "orange";
	//                 cellToHighlight.format.font.bold = true;
	//             })
	//             .then(ctx.sync);
	//     })
	//     .catch(errorHandler);
	// }

	// function displaySelectedCells() {
	//     Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
	//         function (result) {
	//             if (result.status === Office.AsyncResultStatus.Succeeded) {
	//                 showNotification('The selected text is:', '"' + result.value + '"');
	//             } else {
	//                 showNotification('Error', result.error.message);
	//             }
	//         });
	// }

	// // Helper function for treating errors
	// function errorHandler(error) {
	//     // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
	//     showNotification("Error", error);
	//     console.log("Error: " + error);
	//     if (error instanceof OfficeExtension.Error) {
	//         console.log("Debug info: " + JSON.stringify(error.debugInfo));
	//     }
	// }

	// // Helper function for displaying notifications
	// function showNotification(header, content) {
	//     $("#notification-header").text(header);
	//     $("#notification-body").text(content);
	//     messageBanner.showBanner();
	//     messageBanner.toggleExpansion();
	// }
})();
