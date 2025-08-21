/* global clearInterval, console, CustomFunctions, setInterval */

// Import settings from taskpane
import { getEPMSettings } from '../taskpane/taskpane';
import { getApiUrl } from '../utils/url';
import { exportDataSlice as exportDataSliceApiClient } from '../utils/api-client';


/**
 * Exports a data slice from a specified cube with dimension-member pairs
 * @customfunction
 * @param {string} cubeName Name of the cube (e.g., "main")
 * @param {string[]} pairs Individual dimension-member pairs in format "Dimension#Member"
 * @returns {Promise<string>} Result showing all parameters passed
 */
export function exportDataSlice(cubeName: string, ...pairs: string[]): Promise<string | number> {
  try {
    // First, check if EPM settings are configured
    getEPMSettings();
  } catch (error) {
    // If settings are not configured, return a specific error message
    return Promise.resolve("#EPM_SETTINGS_NOT_SET_OR_INVALID!");
  }
  
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
          suppressMissingBlocks: false,
          rows: [
            {
              dimensions: [dimensions[0]],
              members: [[members[0][0]]]  // Change: Wrap in another array
            }
          ],
          columns: [
            {
              dimensions: dimensions.slice(1),
              members: members.slice(1).map(m => [m[0]])
            }
          ]
        }
      };

      // Call our proxy server instead of the Oracle EPM server directly
      exportDataSliceApiClient(cubeName, payload)
        .then(data => {
          // Print the entire response data
          console.log('Full response:', JSON.stringify(data, null, 2));  // Pretty print the full response
          // Parse the response and extract just the data value
          const numericValue = data.rows[0].data[0];
          if (numericValue === "") {
            resolve("#MISSING_BLOCK!");
          } else {
            const num = Number(numericValue);
            resolve(isNaN(num) ? numericValue : num);
          }
        })
        .catch(error => {
          if (error.message) {
            resolve(`#ERROR: ${error.message}!`);
          } else {
            resolve("#SERVER_ERROR!");
          }
        });

    } catch (error) {
      reject(`Error: ${error instanceof Error ? error.message : error}`);
    }
  });
}


/**
 * Automatically refreshes data by detecting the range and metadata
 * @customfunction
 * @param {string} cubeName Name of the cube (e.g., "PROD")
 * @returns {Promise<string>} Status message
 */
// export function refresh_auto(cubeName: string): Promise<string> {
//   return new Promise((resolve, reject) => {
//     Excel.run(async (context) => {
//       try {
//         // Get the active worksheet
//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const usedRange = sheet.getUsedRange();
//         usedRange.load(["values", "rowCount", "columnCount"]);
        
//         await context.sync();
        
//         const range = usedRange.values;
//         console.log("Raw data from Excel range:", JSON.stringify(range, null, 2));

//         // Don't compress indices by filtering; just validate
//         const hasAnyData = range.some(row => row.some(cell => cell !== ""));
//         if (!hasAnyData) {
//           throw new Error("No data found in the worksheet");
//         }

//         // Find the data structure boundaries on the full usedRange
//         const dataStructure = detectDataStructure(range);        // // Create the binding for the data area (relative to usedRange)
//         console.log("Detected data structure (used to build payload):", JSON.stringify(dataStructure, null, 2));
//         // const dataRange = usedRange
//         //   .getCell(dataStructure.rowIndices[0], dataStructure.dataStartCol)
//         //   .getResizedRange(dataStructure.rowCount - 1, dataStructure.columnCount - 1);

//         // Create payload from detected structure
//         const payload: any = {
//           exportPlanningData: false,
//           gridDefinition: {
//             suppressMissingBlocks: false,
//             columns: [{ members: dataStructure.columnLayers }],
//             rows: [{ members: dataStructure.rowMembersTransposed }]
//           }
//         };

//         if (dataStructure.povMembers.length > 0) {
//           payload.gridDefinition.pov = { members: dataStructure.povMembers };
//         }

//         // Use proxy server for API call
//         console.log("Payload being sent:", JSON.stringify(payload, null, 2));
//         const data = await exportDataSliceApiClient(cubeName, payload);

//         // Populate values at exact header intersections (supports gaps)
//         await populateExcelWithDataAtIndices(context, usedRange, data, dataStructure);
//         await context.sync();

//         resolve("Data refreshed successfully");
//       } catch (error) {
//         reject(`Error: ${error.message}`);
//       }
//     });
//   });
// }

/**
 * Helper function to detect the data structure in the sheet
 */
// export function detectDataStructure(range: any[][]) {
//   // Top of header block = first non-empty row anywhere
//   const headerTop = range.findIndex(row => row.some(cell => cell !== ""));
//   if (headerTop === -1) throw new Error("Could not find any content in the sheet");

//   // First non-empty column in that header row
//   const firstHeaderCol = range[headerTop].findIndex(cell => cell !== "");
//   if (firstHeaderCol === -1) throw new Error("Could not find header start");

//   // First data row = first row below headerTop that has any left headers
//   const firstDataRow = range.findIndex((row, idx) =>
//     idx > headerTop && row.slice(0, firstHeaderCol).some(cell => cell !== "")
//   );
//   if (firstDataRow === -1) throw new Error("Could not find data row start (no left row headers found).");

//   const headerBottom = firstDataRow - 1;

//   // Leaf headers indices from the bottom header row (allow gaps)
//   const colIndices: number[] = [];
//   for (let c = firstHeaderCol; c < range[headerBottom].length; c++) {
//     if (range[headerBottom][c] !== "") colIndices.push(c);
//   }
//   if (colIndices.length === 0) throw new Error("No column headers detected.");

//   // All header row indexes (top → bottom)
//   const headerRowIndices = Array.from({ length: headerBottom - headerTop + 1 }, (_, i) => headerTop + i);

//   // Build layered column members across all header rows with forward-fill
//   const columnLayers: string[][] = [];
//   for (const r of headerRowIndices) {
//     const layer: string[] = [];
//     let last = "";
//     for (const c of colIndices) {
//       let v = range[r][c];
//       if (v === "" || v == null) v = last;
//       else last = v;
//       layer.push(v);
//     }
//     columnLayers.push(layer);
//   }

//   // Row members (left headers) and their absolute row indices
//   const rowIndices: number[] = [];
//   const dataRows: string[][] = [];
//   for (let r = firstDataRow; r < range.length; r++) {
//     const left = range[r].slice(0, firstHeaderCol).filter(cell => cell !== "");
//     if (left.length > 0) {
//       rowIndices.push(r);
//       dataRows.push(left);
//     }
//   }
//   if (rowIndices.length === 0) throw new Error("No row members detected.");

//   // Transpose row members to group by dimension
//   const rowMembersTransposed = dataRows[0].map((_, colIndex) =>
//     dataRows.map(row => row[colIndex])
//   );

//   // Left header column indexes (0..firstHeaderCol-1), preserving gaps
//   const leftHeaderColIndices = Array.from({ length: firstHeaderCol }, (_, i) => i);

//   return {
//     headerTop,
//     headerBottom,
//     headerRowIndices,
//     dataStartCol: firstHeaderCol,
//     leftHeaderColIndices,
//     colIndices,
//     rowIndices,
//     rowCount: rowIndices.length,
//     columnCount: colIndices.length,
//     povMembers: headerTop > 0
//       ? range.slice(0, headerTop)
//           .map(row => [row.slice(firstHeaderCol).find(x => x !== "")])
//           .filter(x => x[0])
//       : [],
//     columnLayers,
//     rowMembersTransposed
//   };
// }


// async function populateExcelWithDataAtIndices(
//   context: Excel.RequestContext,
//   usedRange: Excel.Range,
//   responseData: any,
//   dataStructure: {
//     colIndices: number[];
//     rowIndices: number[];
//     headerRowIndices: number[];
//     leftHeaderColIndices: number[];
//     rowMembersTransposed: string[][]; // Add this to the signature
//   }
// ): Promise<void> {
//   const { colIndices, rowIndices, rowMembersTransposed } = dataStructure;

//   // Create a lookup map from the response data for easy access
//   const responseMap = new Map<string, string[]>();
//   responseData.rows.forEach((row: any) => {
//     // Assuming 'headers' in response are the row members that form the key
//     const key = row.headers.join('|').toLowerCase(); // Normalize to lowercase
//     responseMap.set(key, row.data);
//   });

//   // Log the keys from the response map for debugging
//   console.log("Keys from response map:", Array.from(responseMap.keys()));

//   // Build contiguous column segments to avoid filling gaps as blanks
//   const segments: { startIdx: number; endIdx: number; startCol: number }[] = [];
//   let start = 0;
//   for (let i = 1; i <= colIndices.length; i++) {
//     if (i === colIndices.length || colIndices[i] !== colIndices[i - 1] + 1) {
//       segments.push({ startIdx: start, endIdx: i - 1, startCol: colIndices[start] });
//       start = i;
//     }
//   }

//   const rowCount = Math.min(responseData.rows.length, rowIndices.length);
//   for (let r = 0; r < rowCount; r++) {
//     // Construct the key for the current row from the transposed row members
//     const rowKey = rowMembersTransposed.map(dim => dim[r]).join('|').toLowerCase(); // Normalize to lowercase
//     const rowData = responseMap.get(rowKey);
    
//     // Log the key being looked up for debugging
//     console.log(`Looking up key: ${rowKey}, Found: ${!!rowData}`);
    
//     for (const seg of segments) {
//       const len = seg.endIdx - seg.startIdx + 1;
      
//       // If rowData is found, populate the values, otherwise fill with #MISSING_BLOCK
//       const values = rowData 
//         ? Array.from({ length: len }, (_, k) => {
//             const v = rowData[seg.startIdx + k];
//             if (v === null || v === undefined || v.trim() === "") {
//                 return "#MISSING_BLOCK";
//             }
//             const num = Number(v);
//             return isNaN(num) ? v : num;
//           })
//         : new Array(len).fill("#MISSING_BLOCK");

//       const target = usedRange.getCell(rowIndices[r], seg.startCol).getResizedRange(0, len - 1);
//       target.values = [values];
//       target.format.fill.color = "#ccffcc";
//     }
//   }
// }