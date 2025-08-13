/**
 * Response parsing and Excel cell mapping utilities
 */

// ========== Interfaces ==========

export interface ChunkResponse {
  chunkIndex: number;
  chunkId: string;
  status: 'success' | 'error' | 'pending';
  data?: OLAPResponseData;
  error?: string;
  actualCells?: number;
  processingTime?: number;
  metadata?: ResponseMetadata;
}

export interface OLAPResponseData {
  rows: Array<{
    data: string[];
    members?: string[];
    metadata?: { [key: string]: any };
  }>;
  columns?: Array<{
    name: string;
    type?: string;
    format?: string;
  }>;
  metadata?: {
    totalRows: number;
    totalColumns: number;
    dataSource: string;
    queryTime: number;
  };
}

export interface ResponseMetadata {
  sourceConnection: string;
  queryExecutionTime: number;
  dataRetrievalTime: number;
  cacheHit: boolean;
  compressionRatio?: number;
}

export interface CellMapping {
  sourceRow: number;
  sourceCol: number;
  targetRow: number;
  targetCol: number;
  value: any;
  dataType: 'number' | 'string' | 'formula' | 'empty';
  formatting?: CellFormat;
  validation?: CellValidation;
}

export interface CellFormat {
  numberFormat?: string;
  fontColor?: string;
  backgroundColor?: string;
  fontWeight?: 'normal' | 'bold';
  fontStyle?: 'normal' | 'italic';
  textAlign?: 'left' | 'center' | 'right';
  borders?: BorderStyle;
}

export interface BorderStyle {
  top?: string;
  bottom?: string;
  left?: string;
  right?: string;
}

export interface CellValidation {
  isValid: boolean;
  errorMessage?: string;
  warningMessage?: string;
  suggestedValue?: any;
}

export interface RangeGroup {
  startRow: number;
  startCol: number;
  numRows: number;
  numCols: number;
  values: any[][];
  formatting?: CellFormat;
}

export interface MappingResult {
  cellsUpdated: number;
  rangesUpdated: number;
  success: boolean;
  errors: string[];
  warnings: string[];
  performanceMetrics: {
    totalProcessingTime: number;
    excelUpdateTime: number;
    validationTime: number;
  };
}

export interface AssemblyOptions {
  validateIntegrity: boolean;
  handleMissingChunks: 'error' | 'skip' | 'interpolate';
  sortByChunkIndex: boolean;
  mergeStrategy: 'append' | 'overlay' | 'smart';
}

// ========== Main Response Parser Class ==========

export class ResponseParser {
  private assemblyOptions: AssemblyOptions;

  constructor(options?: Partial<AssemblyOptions>) {
    this.assemblyOptions = {
      validateIntegrity: true,
      handleMissingChunks: 'error',
      sortByChunkIndex: true,
      mergeStrategy: 'smart',
      ...options,
    };
  }

  /**
   * Main function to parse and map response data to Excel
   */
  async parseAndMapResponse(
    responses: ChunkResponse[],
    originalStructure: any, // RangeStructure from data-processor
    targetRange: Excel.Range
  ): Promise<MappingResult> {
    const startTime = performance.now();
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      // Step 1: Validate responses
      const validationResult = this.validateResponses(responses);
      if (!validationResult.isValid) {
        errors.push(...validationResult.errors);
        if (errors.length > 0) {
          throw new Error(`Response validation failed: ${errors.join(', ')}`);
        }
      }
      warnings.push(...validationResult.warnings);

      // Step 2: Assemble complete data from chunks
      const validationTime = performance.now();
      const completeData = this.assembleChunks(responses);

      // Step 3: Validate data integrity
      if (this.assemblyOptions.validateIntegrity) {
        const integrityResult = this.validateDataIntegrity(completeData, responses);
        if (!integrityResult.isValid) {
          warnings.push(...integrityResult.warnings);
        }
      }

      // Step 4: Create cell mappings
      const cellMappings = this.mapToExcelCells(completeData, originalStructure);

      // Step 5: Apply to Excel
      const excelUpdateStart = performance.now();
      const result = await this.applyMappingsToExcel(cellMappings, targetRange);
      const excelUpdateTime = performance.now() - excelUpdateStart;

      const totalProcessingTime = performance.now() - startTime;

      return {
        ...result,
        errors: [...errors, ...result.errors],
        warnings: [...warnings, ...result.warnings],
        performanceMetrics: {
          totalProcessingTime,
          excelUpdateTime,
          validationTime: validationTime - startTime,
        },
      };

    } catch (error) {
      return {
        cellsUpdated: 0,
        rangesUpdated: 0,
        success: false,
        errors: [...errors, error.message],
        warnings,
        performanceMetrics: {
          totalProcessingTime: performance.now() - startTime,
          excelUpdateTime: 0,
          validationTime: 0,
        },
      };
    }
  }

  /**
   * Validate chunk responses
   */
  private validateResponses(responses: ChunkResponse[]): { isValid: boolean; errors: string[]; warnings: string[] } {
    const errors: string[] = [];
    const warnings: string[] = [];

    if (!responses || responses.length === 0) {
      errors.push('No responses provided');
      return { isValid: false, errors, warnings };
    }

    // Check for failed chunks
    const failedChunks = responses.filter(r => r.status === 'error');
    if (failedChunks.length > 0) {
      if (this.assemblyOptions.handleMissingChunks === 'error') {
        errors.push(`${failedChunks.length} chunks failed: ${failedChunks.map(c => c.error).join(', ')}`);
      } else {
        warnings.push(`${failedChunks.length} chunks failed but will be handled according to strategy`);
      }
    }

    // Check for missing chunks
    const maxChunkIndex = Math.max(...responses.map(r => r.chunkIndex));
    const expectedChunks = maxChunkIndex + 1;
    const actualChunks = responses.filter(r => r.status === 'success').length;
    
    if (actualChunks < expectedChunks) {
      const missingChunks = expectedChunks - actualChunks;
      if (this.assemblyOptions.handleMissingChunks === 'error') {
        errors.push(`Missing ${missingChunks} chunks`);
      } else {
        warnings.push(`Missing ${missingChunks} chunks, will be handled according to strategy`);
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * Assemble chunks into complete response data
   */
  private assembleChunks(responses: ChunkResponse[]): OLAPResponseData {
    // Filter successful responses
    let successfulResponses = responses.filter(r => r.status === 'success' && r.data);

    // Sort by chunk index if requested
    if (this.assemblyOptions.sortByChunkIndex) {
      successfulResponses.sort((a, b) => a.chunkIndex - b.chunkIndex);
    }

    if (successfulResponses.length === 0) {
      throw new Error('No successful responses to assemble');
    }

    // If only one chunk, return it directly
    if (successfulResponses.length === 1) {
      return successfulResponses[0].data!;
    }

    // Merge multiple chunks
    const assembledData: OLAPResponseData = {
      rows: [],
      columns: successfulResponses[0].data!.columns, // Use first chunk's column definition
      metadata: {
        totalRows: 0,
        totalColumns: successfulResponses[0].data!.columns?.length || 0,
        dataSource: 'assembled',
        queryTime: successfulResponses.reduce((sum, r) => sum + (r.processingTime || 0), 0),
      },
    };

    // Merge rows based on strategy
    switch (this.assemblyOptions.mergeStrategy) {
      case 'append':
        successfulResponses.forEach(response => {
          if (response.data?.rows) {
            assembledData.rows.push(...response.data.rows);
          }
        });
        break;

      case 'overlay':
        // More complex merging logic for overlapping data
        assembledData.rows = this.mergeRowsWithOverlay(successfulResponses);
        break;

      case 'smart':
      default:
        // Intelligent merging based on data patterns
        assembledData.rows = this.mergeRowsSmart(successfulResponses);
        break;
    }

    assembledData.metadata!.totalRows = assembledData.rows.length;

    return assembledData;
  }

  /**
   * Smart merging of rows from multiple chunks
   */
  private mergeRowsSmart(responses: ChunkResponse[]): OLAPResponseData['rows'] {
    const allRows: OLAPResponseData['rows'] = [];
    
    responses.forEach(response => {
      if (response.data?.rows) {
        allRows.push(...response.data.rows);
      }
    });

    // Remove duplicates based on member combinations if available
    const uniqueRows = allRows.filter((row, index, arr) => {
      if (!row.members) return true; // Keep rows without member info
      
      return arr.findIndex(otherRow => 
        otherRow.members && 
        JSON.stringify(otherRow.members) === JSON.stringify(row.members)
      ) === index;
    });

    return uniqueRows;
  }

  /**
   * Overlay merging (for overlapping data scenarios)
   */
  private mergeRowsWithOverlay(responses: ChunkResponse[]): OLAPResponseData['rows'] {
    // For now, use smart merging
    // In future versions, implement true overlay logic
    return this.mergeRowsSmart(responses);
  }

  /**
   * Validate the integrity of assembled data
   */
  private validateDataIntegrity(
    data: OLAPResponseData, 
    originalResponses: ChunkResponse[]
  ): { isValid: boolean; warnings: string[] } {
    const warnings: string[] = [];

    // Check expected vs actual row count
    const expectedCells = originalResponses.reduce((sum, r) => sum + (r.actualCells || 0), 0);
    const actualCells = data.rows.reduce((sum, row) => sum + row.data.length, 0);

    if (expectedCells !== actualCells) {
      warnings.push(`Cell count mismatch: expected ${expectedCells}, got ${actualCells}`);
    }

    // Check for data consistency
    const columnCount = data.columns?.length || (data.rows[0]?.data.length || 0);
    const inconsistentRows = data.rows.filter(row => row.data.length !== columnCount);
    
    if (inconsistentRows.length > 0) {
      warnings.push(`${inconsistentRows.length} rows have inconsistent column counts`);
    }

    return {
      isValid: warnings.length === 0,
      warnings,
    };
  }

  /**
   * Map response data to Excel cell mappings
   */
  private mapToExcelCells(data: OLAPResponseData, originalStructure: any): CellMapping[] {
    const mappings: CellMapping[] = [];
    
    data.rows.forEach((row, rowIndex) => {
      row.data.forEach((value, colIndex) => {
        const mapping: CellMapping = {
          sourceRow: rowIndex,
          sourceCol: colIndex,
          targetRow: originalStructure.dataStartRow + rowIndex,
          targetCol: originalStructure.dataStartCol + colIndex,
          value: this.parseValue(value),
          dataType: this.determineDataType(value),
          validation: this.validateCellValue(value),
        };

        // Add formatting based on data type
        mapping.formatting = this.getDefaultFormatting(mapping.dataType, value);

        mappings.push(mapping);
      });
    });

    return mappings;
  }

  /**
   * Parse value to appropriate type
   */
  private parseValue(value: string): any {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    // Try to parse as number
    const numValue = parseFloat(value);
    if (!isNaN(numValue) && isFinite(numValue)) {
      return numValue;
    }

    // Return as string
    return value.toString();
  }

  /**
   * Determine data type of a value
   */
  private determineDataType(value: any): CellMapping['dataType'] {
    if (value === null || value === undefined || value === '') {
      return 'empty';
    }

    if (typeof value === 'number') {
      return 'number';
    }

    if (typeof value === 'string') {
      // Check if it's a formula
      if (value.startsWith('=')) {
        return 'formula';
      }
      
      // Check if it can be converted to number
      const numValue = parseFloat(value);
      if (!isNaN(numValue) && isFinite(numValue)) {
        return 'number';
      }
    }

    return 'string';
  }

  /**
   * Validate individual cell value
   */
  private validateCellValue(value: any): CellValidation {
    const validation: CellValidation = { isValid: true };

    // Check for common issues
    if (typeof value === 'string') {
      // Check for error indicators
      if (value.toLowerCase().includes('error') || value.toLowerCase().includes('invalid')) {
        validation.isValid = false;
        validation.errorMessage = 'Cell contains error indicator';
      }

      // Check for very large numbers as strings
      if (value.length > 15 && !isNaN(parseFloat(value))) {
        validation.warningMessage = 'Very large number detected, may lose precision';
      }
    }

    return validation;
  }

  /**
   * Get default formatting for data type
   */
  private getDefaultFormatting(dataType: CellMapping['dataType'], value: any): CellFormat {
    const formatting: CellFormat = {};

    switch (dataType) {
      case 'number':
        formatting.numberFormat = '#,##0.00';
        formatting.textAlign = 'right';
        break;
      case 'string':
        formatting.textAlign = 'left';
        break;
      case 'formula':
        formatting.fontStyle = 'italic';
        formatting.textAlign = 'left';
        break;
      case 'empty':
        formatting.backgroundColor = '#f8f8f8';
        break;
    }

    return formatting;
  }

  /**
   * Apply cell mappings to Excel
   */
  private async applyMappingsToExcel(
    mappings: CellMapping[], 
    targetRange: Excel.Range
  ): Promise<MappingResult> {
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      return await Excel.run(async (context) => {
        // Group mappings by contiguous ranges for batch updates
        const rangeGroups = this.groupMappingsByRange(mappings);
        
        for (const group of rangeGroups) {
          try {
            const range = targetRange.getCell(group.startRow, group.startCol)
              .getResizedRange(group.numRows - 1, group.numCols - 1);
            
            // Highlight before updating (light blue to indicate refreshed data)
            range.format.fill.color = "#E8F4FD";
            
            // Set values
            range.values = group.values;
            
            // Apply formatting if needed
            if (group.formatting) {
              this.applyFormattingToRange(range, group.formatting);
            }
          } catch (error) {
            errors.push(`Failed to update range starting at (${group.startRow}, ${group.startCol}): ${error.message}`);
          }
        }
        
        await context.sync();
        
        return {
          cellsUpdated: mappings.length,
          rangesUpdated: rangeGroups.length,
          success: errors.length === 0,
          errors,
          warnings,
          performanceMetrics: {
            totalProcessingTime: 0, // Will be filled by caller
            excelUpdateTime: 0, // Will be filled by caller
            validationTime: 0, // Will be filled by caller
          },
        };
      });
    } catch (error) {
      return {
        cellsUpdated: 0,
        rangesUpdated: 0,
        success: false,
        errors: [...errors, `Excel update failed: ${error.message}`],
        warnings,
        performanceMetrics: {
          totalProcessingTime: 0,
          excelUpdateTime: 0,
          validationTime: 0,
        },
      };
    }
  }

  /**
   * Group mappings into contiguous ranges for efficient updates
   */
  private groupMappingsByRange(mappings: CellMapping[]): RangeGroup[] {
    if (mappings.length === 0) return [];

    // Sort mappings by row then column
    const sortedMappings = [...mappings].sort((a, b) => {
      if (a.targetRow !== b.targetRow) {
        return a.targetRow - b.targetRow;
      }
      return a.targetCol - b.targetCol;
    });

    const groups: RangeGroup[] = [];
    let currentGroup: RangeGroup | null = null;

    for (const mapping of sortedMappings) {
      if (!currentGroup || 
          !this.canAddToGroup(currentGroup, mapping)) {
        // Start new group
        if (currentGroup) {
          groups.push(currentGroup);
        }
        
        currentGroup = {
          startRow: mapping.targetRow,
          startCol: mapping.targetCol,
          numRows: 1,
          numCols: 1,
          values: [[mapping.value]],
          formatting: mapping.formatting,
        };
      } else {
        // Add to existing group
        this.addMappingToGroup(currentGroup, mapping);
      }
    }

    if (currentGroup) {
      groups.push(currentGroup);
    }

    return groups;
  }

  /**
   * Check if a mapping can be added to an existing group
   */
  private canAddToGroup(group: RangeGroup, mapping: CellMapping): boolean {
    // For simplicity, only group mappings in the same row
    return mapping.targetRow === group.startRow && 
           mapping.targetCol === group.startCol + group.numCols;
  }

  /**
   * Add a mapping to an existing group
   */
  private addMappingToGroup(group: RangeGroup, mapping: CellMapping): void {
    if (mapping.targetRow === group.startRow) {
      // Extend current row
      group.values[0].push(mapping.value);
      group.numCols++;
    } else {
      // This should not happen with current grouping logic
      console.warn('Unexpected mapping row in group');
    }
  }

  /**
   * Apply formatting to an Excel range
   */
  private applyFormattingToRange(range: Excel.Range, formatting: CellFormat): void {
    if (formatting.numberFormat) {
      range.numberFormat = [[formatting.numberFormat]];
    }
    
    if (formatting.fontColor) {
      range.format.font.color = formatting.fontColor;
    }
    
    if (formatting.backgroundColor) {
      range.format.fill.color = formatting.backgroundColor;
    }
    
    if (formatting.fontWeight) {
      range.format.font.bold = formatting.fontWeight === 'bold';
    }
    
    if (formatting.fontStyle) {
      range.format.font.italic = formatting.fontStyle === 'italic';
    }
    
    if (formatting.textAlign) {
      const alignment = formatting.textAlign === 'left' ? Excel.HorizontalAlignment.left :
                       formatting.textAlign === 'center' ? Excel.HorizontalAlignment.center :
                       Excel.HorizontalAlignment.right;
      range.format.horizontalAlignment = alignment;
    }
  }

  /**
   * Update assembly options
   */
  updateAssemblyOptions(options: Partial<AssemblyOptions>): void {
    this.assemblyOptions = { ...this.assemblyOptions, ...options };
  }

  /**
   * Get current assembly options
   */
  getAssemblyOptions(): AssemblyOptions {
    return { ...this.assemblyOptions };
  }
}

// ========== Utility Functions ==========

/**
 * Create a response parser with default settings
 */
export function createResponseParser(options?: Partial<AssemblyOptions>): ResponseParser {
  return new ResponseParser(options);
}

/**
 * Quick function to validate response chunks
 */
export function validateResponseChunks(responses: ChunkResponse[]): { isValid: boolean; summary: string } {
  const parser = new ResponseParser();
  const validation = parser['validateResponses'](responses); // Access private method for utility
  
  return {
    isValid: validation.isValid,
    summary: `${responses.length} chunks, ${validation.errors.length} errors, ${validation.warnings.length} warnings`,
  };
}
