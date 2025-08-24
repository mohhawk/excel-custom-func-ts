/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { saveReport } from "../utils/api-client";

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification to the user
 * @param message The message to show
 */
function showNotification(message: string) {
  console.log(message);
  // You can enhance this with actual UI notifications if needed
}

function showExportReportDialog(event: Office.AddinCommands.Event) {
  console.log("Export Report button clicked. Executing showExportReportDialog...");
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();

    const reportName = sheet.name;
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    
    const exportData = {
      name: `${reportName}-${timestamp}`,
      version: "1.0.0",
      changelog: "Direct export",
      includeFormatting: true, // Defaulting to true
    };

    await processExport(exportData);
    showNotification(`Report '${exportData.name}' has been exported.`);
    event.completed();
  }).catch((error) => {
    console.error(error);
    showNotification("Error exporting report.");
    event.completed();
  });
}

function showImportReportDialog(event: Office.AddinCommands.Event) {
    console.log("showImportReportDialog called");
    event.completed();
}

async function processExport(exportData: { name: string, version: string, changelog: string, includeFormatting: boolean }) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.load(["values", "address"]);

        await context.sync();
        
        const report = {
            name: exportData.name,
            version: exportData.version,
            changelog: exportData.changelog,
            data: null as any,
        };

        if (exportData.includeFormatting) {
            range.load(["format/fill", "format/font", "format/borders"]);
            await context.sync();
            
            const cellData = [];
            for (let i = 0; i < range.values.length; i++) {
                const rowData = [];
                for (let j = 0; j < range.values[i].length; j++) {
                    const cell = range.getCell(i, j);
                    cell.load(["values", "format/fill", "format/font", "format/borders"]);
                    await context.sync();
                    rowData.push({
                        value: cell.values[0][0],
                        format: {
                            fill: cell.format.fill,
                            font: cell.format.font,
                            borders: cell.format.borders
                        }
                    });
                }
                cellData.push(rowData);
            }
            report.data = cellData;
        } else {
            report.data = range.values;
        }

        console.log("Report to be sent to backend:", report);
        await saveReport(report);
    });
}


/**
 * Handles the refresh data button click
 */
// async function refreshData(event: Office.AddinCommands.Event) {
//   try {
//     // Show a dialog to get cube name
//     Office.context.ui.displayDialogAsync(
//       'https://localhost:3000/cubename.html',
//       { height: 30, width: 20 },
//       (result) => {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           const dialog = result.value;
//           dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (args) => {
//             // Fix the type issue
//             const messageData = args as { message: string; origin: string; };
//             const cubeName = messageData.message;
//             dialog.close();
            
//             try {
//               // Call the new refresh function
//               await refreshAdhocData(cubeName);
//               showNotification('Data refreshed successfully!');
//             } catch (error) {
//               showNotification(`Error refreshing data: ${error.message}`);
//             }
//           });
//         } else {
//           showNotification('Failed to open cube name dialog');
//         }
//       }
//     );
    
//     event.completed();
//   } catch (error) {
//     console.error('Error:', error);
//     showNotification(`Error: ${error.message}`);
//     event.completed();
//   }
// }

// Register the function
// Office.actions.associate("refreshData", refreshData);
Office.actions.associate("showExportReportDialog", showExportReportDialog);
Office.actions.associate("showImportReportDialog", showImportReportDialog);

// Make functions available on the window object for robustness
(window as any).showExportReportDialog = showExportReportDialog;
(window as any).showImportReportDialog = showImportReportDialog;
