/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { refreshAdhocData } from '../taskpane/taskpane';

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

/**
 * Handles the refresh data button click
 */
async function refreshData(event: Office.AddinCommands.Event) {
  try {
    // Show a dialog to get cube name
    Office.context.ui.displayDialogAsync(
      'https://localhost:3000/cubename.html',
      { height: 30, width: 20 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (args) => {
            // Fix the type issue
            const messageData = args as { message: string; origin: string; };
            const cubeName = messageData.message;
            dialog.close();
            
            try {
              // Call the new refresh function
              await refreshAdhocData(cubeName);
              showNotification('Data refreshed successfully!');
            } catch (error) {
              showNotification(`Error refreshing data: ${error.message}`);
            }
          });
        } else {
          showNotification('Failed to open cube name dialog');
        }
      }
    );
    
    event.completed();
  } catch (error) {
    console.error('Error:', error);
    showNotification(`Error: ${error.message}`);
    event.completed();
  }
}

// Register the function
Office.actions.associate("refreshData", refreshData);
