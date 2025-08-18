/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import {
  getAppConfig,
  isDevEnvironment
} from '../utils/config';
import { getApiUrl } from '../utils/url';
import { detectDataStructure } from '../functions/functions';
import { exportDataSlice } from '../utils/api-client';

interface EPMSettings {
  connectionType: 'hyperion' | 'tinyolap';
  serverUrl: string;
  application: string;
  username?: string;
  password?: string;
  tinyOlapServerUrl?: string;
  tinyOlapApplication?: string;
}

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  
  // Load saved settings
  loadSavedSettings();
  
  // Add event listeners
  document.getElementById("save-settings").onclick = saveSettings;
  document.getElementById("refresh-data").onclick = () => refreshAdhocData('main');
  document.getElementById('connection-type').addEventListener('change', toggleConnectionFields);
});

function toggleConnectionFields() {
  const connectionType = (document.getElementById('connection-type') as HTMLSelectElement).value;
  if (connectionType === 'hyperion') {
    document.getElementById('hyperion-fields').classList.remove('hidden');
    document.getElementById('tinyolap-fields').classList.add('hidden');
  } else {
    document.getElementById('hyperion-fields').classList.add('hidden');
    document.getElementById('tinyolap-fields').classList.remove('hidden');
  }
}

function loadSavedSettings() {
  const savedSettings = localStorage.getItem('epmSettings');
  if (savedSettings) {
    const settings: EPMSettings = JSON.parse(savedSettings);
    (document.getElementById('connection-type') as HTMLSelectElement).value = settings.connectionType || 'hyperion';
    (document.getElementById('server-url') as HTMLInputElement).value = settings.serverUrl || '';
    (document.getElementById('application') as HTMLInputElement).value = settings.application || '';
    (document.getElementById('username') as HTMLInputElement).value = settings.username || '';
    (document.getElementById('password') as HTMLInputElement).value = settings.password || '';
    (document.getElementById('tinyolap-server-url') as HTMLInputElement).value = settings.tinyOlapServerUrl || '';
    (document.getElementById('tinyolap-application') as HTMLInputElement).value = settings.tinyOlapApplication || '';
    toggleConnectionFields();
  }
}

function saveSettings() {
  const connectionType = (document.getElementById('connection-type') as HTMLSelectElement).value as 'hyperion' | 'tinyolap';
  
  const settings: EPMSettings = {
    connectionType,
    serverUrl: (document.getElementById('server-url') as HTMLInputElement).value,
    application: (document.getElementById('application') as HTMLInputElement).value,
    username: (document.getElementById('username') as HTMLInputElement).value,
    password: (document.getElementById('password') as HTMLInputElement).value,
    tinyOlapServerUrl: (document.getElementById('tinyolap-server-url') as HTMLInputElement).value,
    tinyOlapApplication: (document.getElementById('tinyolap-application') as HTMLInputElement).value,
  };

  if (connectionType === 'tinyolap') {
    settings.username = '';
    settings.password = '';
  }

  // Save to localStorage
  localStorage.setItem('epmSettings', JSON.stringify(settings));

  // Show success message
  const statusElement = document.getElementById('settings-status');
  statusElement.classList.remove('hidden');
  setTimeout(() => {
    statusElement.classList.add('hidden');
  }, 3000);
}

// Export settings getter for use in functions.ts
export function getEPMSettings(): EPMSettings {
  const savedSettings = localStorage.getItem('epmSettings');
  if (!savedSettings) {
    throw new Error('EPM settings not configured. Please configure settings in the taskpane.');
  }
  const settings: EPMSettings = JSON.parse(savedSettings);

  if (settings.connectionType === 'tinyolap') {
    settings.serverUrl = `${settings.tinyOlapServerUrl}/api/v1/app/${settings.tinyOlapApplication}/cube/{cube_name}/slice/`;
  }
  
  return settings;
}

/**
 * Shows a notification in the taskpane
 * @param message The message to show
 * @param isError Whether this is an error message
 */
function showNotification(message: string, isError: boolean = false) {
  // Create or update notification element
  let notificationElement = document.getElementById('refresh-notification');
  if (!notificationElement) {
    notificationElement = document.createElement('div');
    notificationElement.id = 'refresh-notification';
    notificationElement.className = 'ms-MessageBar';
    
    // Insert after the refresh button
    const refreshButton = document.getElementById('refresh-data');
    refreshButton.parentNode.insertBefore(notificationElement, refreshButton.nextSibling);
  }
  
  // Update classes and content
  notificationElement.className = isError ? 
    'ms-MessageBar ms-MessageBar--error' : 
    'ms-MessageBar ms-MessageBar--success';
  
  notificationElement.innerHTML = `<span class="ms-MessageBar-text">${message}</span>`;
  notificationElement.style.display = 'block';
  
  // Hide after 5 seconds
  setTimeout(() => {
    notificationElement.style.display = 'none';
  }, 5000);
}

export async function refreshAdhocData(cubeName?: string) {
  try {
    showNotification('Refreshing data...', false);
    
    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["values", "address"]);
      await context.sync();

      const range = selectedRange.values as string[][];
      
      const dataStructure = detectDataStructure(range);

      // Create the payload structure
      const payload: any = {
        exportPlanningData: false,
        gridDefinition: {
          suppressMissingBlocks: false,
          columns: [{ members: dataStructure.columnLayers }],
          rows: [{ members: dataStructure.rowMembersTransposed }]
        }
      };
      
      if (dataStructure.povMembers.length > 0) {
        payload.gridDefinition.pov = {
          members: dataStructure.povMembers
        };
      }


      // Make API call
      const data = await exportDataSlice(cubeName || 'main', payload);
      
      console.log('Payload Sent:', JSON.stringify(payload, null, 2));
      console.log('Server Response:', JSON.stringify(data, null, 2));

      // Map response data back to Excel
      if (data.rows && data.rows.length > 0) {
        // Create a lookup map from the response data for easy access
        const responseMap = new Map<string, string[]>();
        data.rows.forEach((row: any) => {
          const key = row.headers.join('|').toLowerCase(); // Normalize to lowercase
          responseMap.set(key, row.data);
        });

        // Log the keys from the response map for debugging
        console.log("Keys from response map:", Array.from(responseMap.keys()));

        // Create the data matrix for Excel, mapping response data to the correct grid location
        const dataMatrix: (string | number)[][] = [];
        
        // Get the original data rows from the Excel grid to match against the response
        const excelDataRows = dataStructure.rowMembersTransposed[0].map((_, colIndex) => 
          dataStructure.rowMembersTransposed.map(row => row[colIndex])
        );
        
        excelDataRows.forEach(rowMembers => {
          const key = rowMembers.join('|').toLowerCase(); // Normalize to lowercase
          const responseData = responseMap.get(key);
          
          // Log the key being looked up for debugging
          console.log(`Looking up key: ${key}, Found: ${!!responseData}`);

          if (responseData) {
            const rowData: (string | number)[] = responseData.map((value: string | null | undefined) => {
                if (value === null || value === undefined || value.trim() === "") {
                    return "#MISSING_BLOCK";
                }
                const numValue = Number(value);
                return isNaN(numValue) ? value : numValue;
            });
            dataMatrix.push(rowData);
          } else {
            // If no data for this row combination, fill with #MISSING_BLOCK
            const placeholderRow = new Array(dataStructure.columnLayers[0].length).fill("#MISSING_BLOCK");
            dataMatrix.push(placeholderRow);
          }
        });

        // Calculate the range where data should be populated
        // Data starts at firstDataRow and firstHeaderCol
        const numDataRows = dataMatrix.length;
        const numDataCols = dataMatrix[0].length;
        
        // Get the data range and populate it
        const dataRange = selectedRange.getCell(dataStructure.rowIndices[0], dataStructure.colIndices[0])
          .getResizedRange(numDataRows - 1, numDataCols - 1);
        
        dataRange.values = dataMatrix;
        
        await context.sync();
        console.log("Data refreshed successfully");
        
        // Show success notification
        showNotification('Data refreshed successfully!', false);
      } else {
        throw new Error("No data received from server");
      }
    });
  } catch (error) {
    console.error("Error refreshing data:", error);
    // Show error notification
    showNotification(`Error: ${error.message}`, true);
  }
}

import { refresh_auto } from '../functions/functions';

// Attach to window so it can be called from HTML
(window as any).refresh_auto = refresh_auto;
