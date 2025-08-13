/**
 * Data preprocessing and chunking utilities for Excel OLAP integration
 */

// ========== Interfaces ==========

export interface ValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

export interface RangeStructure {
  povMembers: string[][];
  columnHeaders: string[];
  rowHeaders: string[];
  dataStartRow: number;
  dataStartCol: number;
  totalRows: number;
  totalCols: number;
  isEmpty: boolean;
}

export interface DimensionData {
  povDimensions: { [key: string]: string[] };
  columnDimensions: string[];
  rowDimensions: string[];
  memberCombinations: DimensionMember[][];
}

export interface DimensionMember {
  dimension: string;
  member: string;
  level?: number;
  parent?: string;
}

export interface ChunkingStrategy {
  maxPayloadSize: number; // in bytes
  maxCellsPerChunk: number;
  chunkByDimension: boolean;
  preserveStructure: boolean;
}

export interface DataChunk {
  chunkId: string;
  chunkIndex: number;
  totalChunks: number;
  gridDefinition: GridDefinition;
  estimatedCells: number;
  estimatedSize: number; // in bytes
  metadata: ChunkMetadata;
}

export interface ChunkMetadata {
  originalRange: {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };
  dimensionRanges: {
    povStart: number;
    povEnd: number;
    columnStart: number;
    columnEnd: number;
    rowStart: number;
    rowEnd: number;
  };
  processingHints: {
    requiresTransformation: boolean;
    hasFormulas: boolean;
    hasConditionalData: boolean;
  };
}

export interface GridDefinition {
  exportPlanningData: boolean;
  suppressMissingBlocks: boolean;
  pov?: {
    members: string[][];
  };
  columns: Array<{
    dimensions?: string[];
    members: string[][];
  }>;
  rows: Array<{
    dimensions?: string[];
    members: string[][];
  }>;
}

export interface PreprocessedData {
  structure: RangeStructure;
  dimensions: DimensionData;
  chunks: DataChunk[];
  metadata: {
    originalRange: string[][];
    estimatedTotalCells: number;
    estimatedTotalSize: number;
    chunkCount: number;
    processingTime: number;
  };
}

// ========== Main Data Processor Class ==========

export class DataPreprocessor {
  private chunkingStrategy: ChunkingStrategy;

  constructor(strategy?: Partial<ChunkingStrategy>) {
    this.chunkingStrategy = {
      maxPayloadSize: 1024 * 1024, // 1MB default
      maxCellsPerChunk: 10000,
      chunkByDimension: true,
      preserveStructure: true,
      ...strategy,
    };
  }

  /**
   * Main preprocessing function that orchestrates the entire process
   */
  async preprocessData(range: string[][]): Promise<PreprocessedData> {
    const startTime = performance.now();

    try {
      // Step 1: Validate the input data
      const validation = this.validateRange(range);
      if (!validation.isValid) {
        throw new Error(`Data validation failed: ${validation.errors.join(', ')}`);
      }

      // Step 2: Clean the data
      const cleanedRange = this.cleanEmptyRows(range);

      // Step 3: Identify the structure
      const structure = this.identifyStructure(cleanedRange);
      if (structure.isEmpty) {
        throw new Error('No valid data structure found in the range');
      }

      // Step 4: Extract dimensions
      const dimensions = this.extractDimensions(cleanedRange, structure);

      // Step 5: Create chunks
      const chunks = await this.createChunks(dimensions, structure);

      const processingTime = performance.now() - startTime;

      return {
        structure,
        dimensions,
        chunks,
        metadata: {
          originalRange: range,
          estimatedTotalCells: this.estimateTotalCells(chunks),
          estimatedTotalSize: this.estimateTotalSize(chunks),
          chunkCount: chunks.length,
          processingTime,
        },
      };
    } catch (error) {
      throw new Error(`Preprocessing failed: ${error.message}`);
    }
  }

  /**
   * Validate the input range for basic requirements
   */
  validateRange(range: string[][]): ValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Check if range exists
    if (!range || range.length === 0) {
      errors.push('Range is empty or undefined');
      return { isValid: false, errors, warnings };
    }

    // Check for minimum dimensions
    if (range.length < 2) {
      errors.push('Range must have at least 2 rows');
    }

    if (range[0].length < 2) {
      errors.push('Range must have at least 2 columns');
    }

    // Check for excessive size
    const totalCells = range.length * range[0].length;
    if (totalCells > 100000) {
      warnings.push(`Large range detected (${totalCells} cells). Consider using smaller ranges for better performance.`);
    }

    // Check for inconsistent row lengths
    const firstRowLength = range[0].length;
    const inconsistentRows = range.findIndex((row, index) => 
      index > 0 && row.length !== firstRowLength
    );
    
    if (inconsistentRows !== -1) {
      warnings.push(`Row ${inconsistentRows + 1} has different length than first row`);
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * Remove completely empty rows and normalize the data
   */
  cleanEmptyRows(range: string[][]): string[][] {
    return range.filter(row => 
      row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== "")
    ).map(row => 
      row.map(cell => cell === null || cell === undefined ? "" : cell.toString().trim())
    );
  }

  /**
   * Identify the structure of the range (POV, headers, data area)
   */
  identifyStructure(range: string[][]): RangeStructure {
    if (!range || range.length === 0) {
      return {
        povMembers: [],
        columnHeaders: [],
        rowHeaders: [],
        dataStartRow: 0,
        dataStartCol: 0,
        totalRows: 0,
        totalCols: 0,
        isEmpty: true,
      };
    }

    // Find first non-empty column in first row (where POV members start)
    const firstHeaderCol = range[0].findIndex(cell => cell !== "");
    if (firstHeaderCol === -1) {
      throw new Error("Could not find header start");
    }

    // Find first row with content in column A (where data rows start)
    const firstDataRow = range.findIndex(row => row[0] !== "");
    if (firstDataRow === -1) {
      throw new Error("Could not find data row start");
    }

    // Extract POV members (everything before data rows, excluding the last header row)
    const povMembers = range.slice(0, Math.max(0, firstDataRow - 1))
      .map(row => row.slice(firstHeaderCol).filter(cell => cell !== ""))
      .filter(members => members.length > 0);

    // Get column headers (from the last header row)
    const columnHeaders = firstDataRow > 0 
      ? range[firstDataRow - 1].slice(firstHeaderCol).filter(cell => cell !== "")
      : [];

    // Get row headers (first column of data rows)
    const rowHeaders = range.slice(firstDataRow)
      .map(row => row[0])
      .filter(header => header !== "");

    return {
      povMembers,
      columnHeaders,
      rowHeaders,
      dataStartRow: firstDataRow,
      dataStartCol: firstHeaderCol,
      totalRows: range.length,
      totalCols: range[0]?.length || 0,
      isEmpty: false,
    };
  }

  /**
   * Extract dimension data from the structured range
   */
  extractDimensions(range: string[][], structure: RangeStructure): DimensionData {
    const povDimensions: { [key: string]: string[] } = {};
    
    // Process POV members
    structure.povMembers.forEach((members, index) => {
      const dimensionName = `POV_${index + 1}`;
      povDimensions[dimensionName] = members;
    });

    // Extract row dimension data
    const dataRows = range.slice(structure.dataStartRow)
      .map(row => row.slice(0, structure.dataStartCol).filter(cell => cell !== ""))
      .filter(row => row.length > 0);

    // Transpose data rows to group by dimension
    const rowDimensions: string[] = [];
    const memberCombinations: DimensionMember[][] = [];

    if (dataRows.length > 0) {
      const maxDimensions = Math.max(...dataRows.map(row => row.length));
      
      for (let dimIndex = 0; dimIndex < maxDimensions; dimIndex++) {
        const dimensionName = `RowDim_${dimIndex + 1}`;
        rowDimensions.push(dimensionName);
        
        const members = dataRows.map(row => row[dimIndex] || "").filter(member => member !== "");
        const uniqueMembers = [...new Set(members)];
        
        memberCombinations.push(
          uniqueMembers.map(member => ({
            dimension: dimensionName,
            member: member,
          }))
        );
      }
    }

    return {
      povDimensions,
      columnDimensions: structure.columnHeaders,
      rowDimensions,
      memberCombinations,
    };
  }

  /**
   * Create chunks based on the chunking strategy
   */
  private async createChunks(dimensions: DimensionData, structure: RangeStructure): Promise<DataChunk[]> {
    const chunks: DataChunk[] = [];
    
    // Calculate total estimated cells
    const totalCells = this.estimateCellsFromDimensions(dimensions);
    
    if (totalCells <= this.chunkingStrategy.maxCellsPerChunk) {
      // Single chunk
      const chunk = this.createSingleChunk(dimensions, structure, 0, 1);
      chunks.push(chunk);
    } else {
      // Multiple chunks needed
      const chunkCount = Math.ceil(totalCells / this.chunkingStrategy.maxCellsPerChunk);
      
      if (this.chunkingStrategy.chunkByDimension) {
        chunks.push(...this.createDimensionBasedChunks(dimensions, structure, chunkCount));
      } else {
        chunks.push(...this.createSizeBasedChunks(dimensions, structure, chunkCount));
      }
    }

    return chunks;
  }

  /**
   * Create a single chunk containing all data
   */
  private createSingleChunk(
    dimensions: DimensionData, 
    structure: RangeStructure, 
    chunkIndex: number, 
    totalChunks: number
  ): DataChunk {
    const chunkId = `chunk_${chunkIndex}_${Date.now()}`;
    
    const gridDefinition: GridDefinition = {
      exportPlanningData: false,
      suppressMissingBlocks: true,
      pov: Object.keys(dimensions.povDimensions).length > 0 ? {
        members: Object.values(dimensions.povDimensions)
      } : undefined,
      columns: [{
        members: [dimensions.columnDimensions]
      }],
      rows: dimensions.memberCombinations.length > 0 ? [{
        members: dimensions.memberCombinations.map(combo => combo.map(dm => dm.member))
      }] : [{
        members: [[]]
      }]
    };

    return {
      chunkId,
      chunkIndex,
      totalChunks,
      gridDefinition,
      estimatedCells: this.estimateCellsFromDimensions(dimensions),
      estimatedSize: this.estimateChunkSize(gridDefinition),
      metadata: {
        originalRange: {
          startRow: 0,
          startCol: 0,
          endRow: structure.totalRows - 1,
          endCol: structure.totalCols - 1,
        },
        dimensionRanges: {
          povStart: 0,
          povEnd: structure.dataStartRow - 1,
          columnStart: structure.dataStartCol,
          columnEnd: structure.totalCols - 1,
          rowStart: structure.dataStartRow,
          rowEnd: structure.totalRows - 1,
        },
        processingHints: {
          requiresTransformation: false,
          hasFormulas: false,
          hasConditionalData: false,
        },
      },
    };
  }

  /**
   * Create chunks based on dimension boundaries
   */
  private createDimensionBasedChunks(
    dimensions: DimensionData, 
    structure: RangeStructure, 
    targetChunkCount: number
  ): DataChunk[] {
    const chunks: DataChunk[] = [];
    
    // Split by row dimensions if we have multiple member combinations
    if (dimensions.memberCombinations.length > 0) {
      const membersPerChunk = Math.ceil(dimensions.memberCombinations[0].length / targetChunkCount);
      
      for (let i = 0; i < targetChunkCount; i++) {
        const startIndex = i * membersPerChunk;
        const endIndex = Math.min(startIndex + membersPerChunk, dimensions.memberCombinations[0].length);
        
        if (startIndex < dimensions.memberCombinations[0].length) {
          const chunkDimensions = {
            ...dimensions,
            memberCombinations: dimensions.memberCombinations.map(combo => 
              combo.slice(startIndex, endIndex)
            )
          };
          
          const chunk = this.createSingleChunk(chunkDimensions, structure, i, targetChunkCount);
          chunks.push(chunk);
        }
      }
    } else {
      // Fallback to single chunk if no member combinations
      const chunk = this.createSingleChunk(dimensions, structure, 0, 1);
      chunks.push(chunk);
    }
    
    return chunks;
  }

  /**
   * Create chunks based on size limits
   */
  private createSizeBasedChunks(
    dimensions: DimensionData, 
    structure: RangeStructure, 
    targetChunkCount: number
  ): DataChunk[] {
    // For now, fall back to dimension-based chunking
    // In future versions, we can implement more sophisticated size-based splitting
    return this.createDimensionBasedChunks(dimensions, structure, targetChunkCount);
  }

  /**
   * Estimate the number of cells from dimension data
   */
  private estimateCellsFromDimensions(dimensions: DimensionData): number {
    const columnCount = dimensions.columnDimensions.length || 1;
    const rowCount = dimensions.memberCombinations.reduce((total, combo) => total + combo.length, 0) || 1;
    return columnCount * rowCount;
  }

  /**
   * Estimate the total cells across all chunks
   */
  private estimateTotalCells(chunks: DataChunk[]): number {
    return chunks.reduce((total, chunk) => total + chunk.estimatedCells, 0);
  }

  /**
   * Estimate the total size across all chunks
   */
  private estimateTotalSize(chunks: DataChunk[]): number {
    return chunks.reduce((total, chunk) => total + chunk.estimatedSize, 0);
  }

  /**
   * Estimate the size of a chunk in bytes
   */
  private estimateChunkSize(gridDefinition: GridDefinition): number {
    // Rough estimation: JSON.stringify length * 2 (for unicode) + overhead
    const jsonString = JSON.stringify(gridDefinition);
    return jsonString.length * 2 + 1024; // 1KB overhead
  }

  /**
   * Update chunking strategy
   */
  updateChunkingStrategy(strategy: Partial<ChunkingStrategy>): void {
    this.chunkingStrategy = { ...this.chunkingStrategy, ...strategy };
  }

  /**
   * Get current chunking strategy
   */
  getChunkingStrategy(): ChunkingStrategy {
    return { ...this.chunkingStrategy };
  }
}

// ========== Utility Functions ==========

/**
 * Create a data processor with default settings
 */
export function createDataProcessor(strategy?: Partial<ChunkingStrategy>): DataPreprocessor {
  return new DataPreprocessor(strategy);
}

/**
 * Quick validation function for external use
 */
export function validateExcelRange(range: string[][]): ValidationResult {
  const processor = new DataPreprocessor();
  return processor.validateRange(range);
}

/**
 * Quick estimation function for cost calculation
 */
export async function estimateProcessingCost(
  range: string[][], 
  costPerCell: number = 0.001
): Promise<{ estimatedCells: number; estimatedCost: number; chunkCount: number }> {
  try {
    const processor = new DataPreprocessor();
    const processed = await processor.preprocessData(range);
    
    return {
      estimatedCells: processed.metadata.estimatedTotalCells,
      estimatedCost: processed.metadata.estimatedTotalCells * costPerCell,
      chunkCount: processed.metadata.chunkCount,
    };
  } catch (error) {
    throw new Error(`Cost estimation failed: ${error.message}`);
  }
}
