/* global clearInterval, console, CustomFunctions, setInterval */

// Planning API Configuration
// Remove the hardcoded constants at the top
// const SERVER_URL = "https://epmconfluence-test-epmconfluence.epm.us-phoenix-1.ocs.oraclecloud.com";
// const APPLICATION = "CONFPLAN";
// const PLANNING_API = `/HyperionPlanning/rest/v3/applications/${APPLICATION}`;
// const USERNAME = "itsupport@jirventures.com";
// const PASSWORD = "Oracle2025@101p!";

// Import settings from taskpane
import { getEPMSettings } from '../taskpane/taskpane';

// Batch processing setup
let _batch = [];
let _isBatchedRequestScheduled = false;

/**
 * Makes the actual API call to the remote service
 * @param requestBatch The batch of requests to process
 */
async function _fetchFromRemoteService(requestBatch) {
  const settings = getEPMSettings();
  const planningApi = `/HyperionPlanning/rest/v3/applications/${settings.application}`;
  const url = `${settings.serverUrl}${planningApi}/plantypes/${requestBatch.cubeName}/exportdataslice`;
  
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Basic ' + btoa(`${settings.username}:${settings.password}`),
      },
      body: JSON.stringify(requestBatch.payload)
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    return [{ result: JSON.stringify(data) }];
  } catch (error) {
    return [{ error: error.message }];
  }
}

/**
 * Process the batch of requests
 */
function _makeRemoteRequest() {
  try {
    const batchCopy = _batch.splice(0, _batch.length);
    _isBatchedRequestScheduled = false;

    // Process the first request in the batch (we'll enhance this for multiple requests later)
    const request = batchCopy[0];
    
    _fetchFromRemoteService(request)
      .then((responseBatch) => {
        responseBatch.forEach((response, index) => {
          if (response.error) {
            batchCopy[index].reject(new Error(response.error));
          } else {
            batchCopy[index].resolve(response.result);
          }
        });
      })
      .catch((error) => {
        batchCopy.forEach((item) => {
          item.reject(error);
        });
      });
  } catch (error) {
    console.error("Batch processing error:", error);
  }
}

/**
 * Exports a data slice from a specified cube with dimension-member pairs
 * @customfunction
 * @param {string} cubeName Name of the cube (e.g., "PROD")
 * @param {string[]} pairs Individual dimension-member pairs in format "Dimension#Member"
 * @returns {Promise<string>} Result showing all parameters passed
 */
export function exportDataSlice(cubeName: string, ...pairs: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    try {
      // Remove the last element which is the invocation context
      pairs.pop();
      const flatPairs = pairs.flat();

      const dimensions: string[] = [];
      const members: string[][] = [];

      // Extract dimensions and members from pairs
      flatPairs.forEach((pair, index) => {
        if (typeof pair !== 'string') {
          throw new Error(`Invalid pair at position ${index}: expected string but got ${typeof pair}`);
        }
        
        if (!pair.includes('#')) {
          throw new Error(`Invalid pair format at position ${index}: ${pair}. Expected format: "Dimension#Member"`);
        }

        const [dimension, member] = pair.split('#');
        dimensions.push(dimension);
        const cleanMember = member.replace(/"/g, '');
        members.push([cleanMember]);
      });

      // Construct the payload
      const payload = {
        exportPlanningData: false,
        gridDefinition: {
          suppressMissingBlocks: true,
          rows: [
            {
              dimensions: [dimensions[0]],
              members: [[members[0][0]]]  // Change: Wrap in another array
            }
          ],
          columns: [
            {
              dimensions: dimensions.slice(1),
              members: members.slice(1).map(m => [m[0]])  // Change: Wrap each member in an array
            }
          ]
        }
      };

      // Call our proxy server instead of the Oracle EPM server directly
      fetch('http://localhost:3001/api/exportDataSlice', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          cubeName,
          payload
        })
      })
      .then(response => {
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
      })
      .then(data => {
        // Print the entire response data
        console.log('Full response:', JSON.stringify(data, null, 2));  // Pretty print the full response
        // Parse the response and extract just the data value
        const numericValue = data.rows[0].data[0];
        resolve(numericValue.toString());  // Convert to string since the function returns Promise<string>
      })
      .catch(error => {
        reject(`Error: ${error.message}`);
      });

    } catch (error) {
      reject(`Error: ${error.message}`);
    }
  });
}


/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(
  incrementBy: number,
  invocation: CustomFunctions.StreamingInvocation<number>
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Generates a Planning API payload from Excel range data
 * @customfunction GENERATEPAYLOAD
 * @param {string[][]} range The Excel range containing the data
 * @returns {string} The generated payload as a JSON string
 */
export function generatePayload(range: string[][]): string {
  try {
    // Remove completely empty rows
    const nonEmptyRows = range.filter(row => row.some(cell => cell !== ""));
    
    if (nonEmptyRows.length === 0) {
      throw new Error("No data found in the range");
    }

    // Find first non-empty column in first row (where POV members start)
    const firstHeaderCol = nonEmptyRows[0].findIndex(cell => cell !== "");
    if (firstHeaderCol === -1) {
      throw new Error("Could not find header start");
    }

    // Find first row with content in column A (where data rows start)
    const firstDataRow = nonEmptyRows.findIndex(row => row[0] !== "");
    if (firstDataRow === -1) {
      throw new Error("Could not find data row start");
    }

    // Extract POV members (from first non-empty column to end, up to data rows)
    const povRows = nonEmptyRows.slice(0, firstDataRow - 1); // Exclude month row
    const povMembers = povRows.map(row => 
      [row.slice(firstHeaderCol).filter(cell => cell !== "")[0]] // Wrap each member in its own array
    ).filter(member => member[0]); // Remove any empty members

    // Get months (last POV row)
    const monthRow = nonEmptyRows[firstDataRow - 1];
    const months = monthRow.slice(firstHeaderCol).filter(cell => cell !== "");

    // Extract data rows
    const dataRows = nonEmptyRows.slice(firstDataRow)
      .map(row => row.slice(0, firstHeaderCol).filter(cell => cell !== ""))
      .filter(row => row.length > 0);

    // Transpose data rows to group by dimension (with correct nesting)
    const transposedRows = {
      members: dataRows[0].map((_, colIndex) => 
        dataRows.map(row => row[colIndex]) // This creates a single array of members for each dimension
      )
    };

    // Create the payload structure
    const payload = {
      exportPlanningData: false,
      gridDefinition: {
        suppressMissingBlocks: true,
        pov: {
          members: povMembers
        },
        columns: [
          {
            members: [months]
          }
        ],
        rows: [transposedRows] // Single array containing the transposed rows object
      }
    };

    // Log debugging information
    console.log('Grid Structure:', {
      firstHeaderCol,
      firstDataRow,
      povMembers,
      months,
      transposedRows,
      dataRowsCount: dataRows.length
    });

    // Log the payload to console
    console.log('Generated Payload:', JSON.stringify(payload, null, 2));

    return JSON.stringify(payload);
  } catch (error) {
    console.error('Error generating payload:', error);
    return JSON.stringify({ error: error.message });
  }
}
