/**
 * Advanced Excel operations tools
 * Features for reading, writing, and analyzing Excel files
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../../graph/client.js";
import { jsonTextResponse, toolErrorResponse } from "../../graph/contracts.js";
import { DriveItem, WorkbookSession } from "../../graph/models.js";
import { getDriveRootEndpoint } from "../../graph/resource-resolver.js";
import { resolveDriveTargetContext } from "../../sharepoint/site-resolver.js";

// Tool 1: Excel workbook operations
export const excelOperations: Tool = {
  name: "excel_operations",
  description: "Perform operations on Excel workbooks (read, write, formulas)",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Excel file item ID",
      },
      itemPath: {
        type: "string",
        description: "Alternative: Excel file path",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      operation: {
        type: "string",
        enum: [
          "read_range",
          "write_range",
          "add_worksheet",
          "list_worksheets",
          "get_formulas",
          "set_formulas",
          "create_table",
          "create_chart",
        ],
        description: "Operation to perform",
      },
      worksheet: {
        type: "string",
        description: "Worksheet name",
        default: "Sheet1",
      },
      range: {
        type: "string",
        description: 'Cell range (e.g., "A1:C10")',
      },
      values: {
        type: "array",
        items: {
          type: "array",
          items: { type: "string", description: "Cell value" },
        },
        description: "Values to write (2D array)",
      },
      formulas: {
        type: "array",
        items: { type: "array", items: { type: "string" } },
        description: "Formulas to set (2D array of formula strings)",
      },
      tableName: {
        type: "string",
        description: "Name for new table",
      },
      hasHeaders: {
        type: "boolean",
        description: "Whether table has headers",
        default: true,
      },
      chartType: {
        type: "string",
        enum: [
          "ColumnClustered",
          "LineMarkers",
          "PieExploded",
          "BarClustered",
          "Area",
          "XYScatterSmooth",
        ],
        description: "Chart type to create",
      },
      useSession: {
        type: "boolean",
        description: "Use persistent session for better performance",
        default: false,
      },
    },
    required: ["operation"],
  },
};

export async function handleExcelOperations(args: any) {
  try {
    const client = getGraphClient();
    const {
      itemId,
      itemPath,
      operation,
      worksheet = "Sheet1",
      range,
      values,
      formulas,
      tableName,
      hasHeaders = true,
      chartType,
      useSession = false,
    } = args;
    const { siteId, driveId } = await resolveDriveTargetContext(
      { site: args.site, siteId: args.siteId, siteUrl: args.siteUrl, driveId: args.driveId },
      client,
    );

    const driveRoot = getDriveRootEndpoint({ siteId, driveId });

    // Get item ID if path is provided
    let actualItemId = itemId;
    if (!actualItemId && itemPath) {
      const itemEndpoint = `${driveRoot}/root:/${itemPath}`;

      const itemResponse = await client.get<DriveItem>(itemEndpoint);
      if (itemResponse.success && itemResponse.data) {
        actualItemId = itemResponse.data.id;
      } else {
        throw new Error("Failed to resolve Excel file path");
      }
    }

    if (!actualItemId) {
      throw new Error("Either itemId or itemPath is required");
    }

    // Base endpoint for workbook operations
    const workbookBase = `${driveRoot}/items/${actualItemId}/workbook`;

    // Create session if requested
    let sessionId: string | undefined;
    if (useSession) {
      const sessionResponse = await client.post<WorkbookSession>(
        `${workbookBase}/createSession`,
        {
          persistChanges: true,
        },
      );

      if (sessionResponse.success && sessionResponse.data) {
        sessionId = sessionResponse.data.id;
      }
    }

    // Add session header if available
    const headers: Record<string, string> | undefined = sessionId
      ? { "workbook-session-id": sessionId }
      : undefined;

    try {
      switch (operation) {
        case "read_range": {
          if (!range) {
            throw new Error("Range is required for read_range operation");
          }

          const endpoint = `${workbookBase}/worksheets/${worksheet}/range(address='${range}')`;
          const response = await client.get<any>(endpoint, {}, { headers });

          if (response.success && response.data) {
            const rangeData = response.data;

            return jsonTextResponse({
              operation: "read_range",
              worksheet,
              range,
              address: rangeData.address,
              rowCount: rangeData.rowCount,
              columnCount: rangeData.columnCount,
              values: rangeData.values,
              formulas: rangeData.formulas,
              text: rangeData.text,
              numberFormat: rangeData.numberFormat,
            });
          }
          break;
        }

        case "write_range": {
          if (!range || !values) {
            throw new Error(
              "Range and values are required for write_range operation",
            );
          }

          const endpoint = `${workbookBase}/worksheets/${worksheet}/range(address='${range}')`;
          const response = await client.patch<any>(
            endpoint,
            { values },
            { headers },
          );

          if (response.success && response.data) {
            return jsonTextResponse({
              operation: "write_range",
              success: true,
              worksheet,
              range,
              rowsWritten: values.length,
              columnsWritten: values[0]?.length || 0,
              address: response.data.address,
            });
          }
          break;
        }

        case "add_worksheet": {
          if (!worksheet) {
            throw new Error(
              "Worksheet name is required for add_worksheet operation",
            );
          }

          const endpoint = `${workbookBase}/worksheets/add`;
          const response = await client.post<any>(
            endpoint,
            { name: worksheet },
            { headers },
          );

          if (response.success && response.data) {
            return jsonTextResponse({
              operation: "add_worksheet",
              success: true,
              worksheet: {
                id: response.data.id,
                name: response.data.name,
                position: response.data.position,
                visibility: response.data.visibility,
              },
            });
          }
          break;
        }

        case "list_worksheets": {
          const endpoint = `${workbookBase}/worksheets`;
          const response = await client.get<any>(endpoint, {}, { headers });

          if (response.success && response.data) {
            const worksheets = (response.data as any).value || [];

            return jsonTextResponse({
              operation: "list_worksheets",
              count: worksheets.length,
              worksheets: worksheets.map((ws: any) => ({
                id: ws.id,
                name: ws.name,
                position: ws.position,
                visibility: ws.visibility,
              })),
            });
          }
          break;
        }

        case "get_formulas": {
          if (!range) {
            throw new Error("Range is required for get_formulas operation");
          }

          const endpoint = `${workbookBase}/worksheets/${worksheet}/range(address='${range}')`;
          const response = await client.get<any>(
            endpoint,
            {
              $select: "formulas,address,rowCount,columnCount",
            },
            { headers },
          );

          if (response.success && response.data) {
            return jsonTextResponse({
              operation: "get_formulas",
              worksheet,
              range,
              address: response.data.address,
              rowCount: response.data.rowCount,
              columnCount: response.data.columnCount,
              formulas: response.data.formulas,
            });
          }
          break;
        }

        case "set_formulas": {
          if (!range || !formulas) {
            throw new Error(
              "Range and formulas are required for set_formulas operation",
            );
          }

          const endpoint = `${workbookBase}/worksheets/${worksheet}/range(address='${range}')`;
          const response = await client.patch<any>(
            endpoint,
            { formulas },
            { headers },
          );

          if (response.success && response.data) {
            return jsonTextResponse({
              operation: "set_formulas",
              success: true,
              worksheet,
              range,
              formulasSet: formulas.length * (formulas[0]?.length || 0),
              address: response.data.address,
            });
          }
          break;
        }

        case "create_table": {
          if (!range || !tableName) {
            throw new Error(
              "Range and tableName are required for create_table operation",
            );
          }

          const endpoint = `${workbookBase}/worksheets/${worksheet}/tables/add`;
          const response = await client.post<any>(
            endpoint,
            {
              address: `${worksheet}!${range}`,
              hasHeaders,
            },
            { headers },
          );

          if (response.success && response.data) {
            const table = response.data;

            // Rename the table
            if (tableName !== table.name) {
              const renameEndpoint = `${workbookBase}/tables/${table.id}`;
              await client.patch(
                renameEndpoint,
                { name: tableName },
                { headers },
              );
            }

            return jsonTextResponse({
              operation: "create_table",
              success: true,
              table: {
                id: table.id,
                name: tableName,
                showHeaders: table.showHeaders,
                showTotals: table.showTotals,
                style: table.style,
              },
            });
          }
          break;
        }

        case "create_chart": {
          if (!range || !chartType) {
            throw new Error(
              "Range and chartType are required for create_chart operation",
            );
          }

          const endpoint = `${workbookBase}/worksheets/${worksheet}/charts/add`;
          const response = await client.post<any>(
            endpoint,
            {
              type: chartType,
              sourceData: `${worksheet}!${range}`,
              seriesBy: "auto",
            },
            { headers },
          );

          if (response.success && response.data) {
            const chart = response.data;

            return jsonTextResponse({
              operation: "create_chart",
              success: true,
              chart: {
                id: chart.id,
                name: chart.name,
                type: chartType,
                height: chart.height,
                width: chart.width,
                top: chart.top,
                left: chart.left,
              },
            });
          }
          break;
        }

        default:
          throw new Error(`Invalid operation: ${operation}`);
      }

      throw new Error(`Failed to perform operation: ${operation}`);
    } finally {
      // Close session if it was created
      if (sessionId) {
        try {
          await client.post(
            `${workbookBase}/closeSession`,
            {},
            { headers: { "workbook-session-id": sessionId } },
          );
        } catch (error) {
          // Session close error is not critical
        }
      }
    }
  } catch (error) {
    return toolErrorResponse("excel_operations", error);
  }
}

// Tool 2: Excel data analysis
export const excelAnalysis: Tool = {
  name: "excel_analysis",
  description:
    "Analyze data in Excel workbooks (statistics, pivot tables, data validation)",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Excel file item ID",
      },
      itemPath: {
        type: "string",
        description: "Alternative: Excel file path",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      analysisType: {
        type: "string",
        enum: [
          "statistics",
          "pivot_summary",
          "data_validation",
          "named_ranges",
          "used_range",
        ],
        description: "Type of analysis to perform",
      },
      worksheet: {
        type: "string",
        description: "Worksheet name",
        default: "Sheet1",
      },
      range: {
        type: "string",
        description:
          "Cell range to analyze (optional, uses used range if not specified)",
      },
    },
    required: ["analysisType"],
  },
};

export async function handleExcelAnalysis(args: any) {
  try {
    const client = getGraphClient();
    const {
      itemId,
      itemPath,
      analysisType,
      worksheet = "Sheet1",
      range,
    } = args;
    const { siteId, driveId } = await resolveDriveTargetContext(
      { site: args.site, siteId: args.siteId, siteUrl: args.siteUrl, driveId: args.driveId },
      client,
    );

    const driveRoot = getDriveRootEndpoint({ siteId, driveId });

    // Get item ID if path is provided
    let actualItemId = itemId;
    if (!actualItemId && itemPath) {
      const itemEndpoint = `${driveRoot}/root:/${itemPath}`;

      const itemResponse = await client.get<DriveItem>(itemEndpoint);
      if (itemResponse.success && itemResponse.data) {
        actualItemId = itemResponse.data.id;
      } else {
        throw new Error("Failed to resolve Excel file path");
      }
    }

    if (!actualItemId) {
      throw new Error("Either itemId or itemPath is required");
    }

    // Base endpoint for workbook operations
    const workbookBase = `${driveRoot}/items/${actualItemId}/workbook`;

    switch (analysisType) {
      case "statistics": {
        // Get the range to analyze
        const rangeEndpoint = range
          ? `${workbookBase}/worksheets/${worksheet}/range(address='${range}')`
          : `${workbookBase}/worksheets/${worksheet}/usedRange`;

        const rangeResponse = await client.get<any>(rangeEndpoint);

        if (rangeResponse.success && rangeResponse.data) {
          const data = rangeResponse.data;
          const values = data.values || [];

          // Calculate statistics for numeric columns
          const stats: any[] = [];

          if (values.length > 0) {
            const columnCount = values[0].length;

            for (let col = 0; col < columnCount; col++) {
              const columnValues = values
                .map((row: any[]) => row[col])
                .filter((val: any) => typeof val === "number");

              if (columnValues.length > 0) {
                const sum = columnValues.reduce(
                  (a: number, b: number) => a + b,
                  0,
                );
                const mean = sum / columnValues.length;
                const sortedValues = [...columnValues].sort((a, b) => a - b);
                const median =
                  sortedValues[Math.floor(sortedValues.length / 2)];
                const min = Math.min(...columnValues);
                const max = Math.max(...columnValues);
                const variance =
                  columnValues.reduce(
                    (acc: number, val: number) => acc + Math.pow(val - mean, 2),
                    0,
                  ) / columnValues.length;
                const stdDev = Math.sqrt(variance);

                stats.push({
                  column: col,
                  columnLetter: String.fromCharCode(65 + col),
                  count: columnValues.length,
                  sum,
                  mean: Math.round(mean * 100) / 100,
                  median,
                  min,
                  max,
                  stdDev: Math.round(stdDev * 100) / 100,
                });
              }
            }
          }

          return jsonTextResponse({
            analysisType: "statistics",
            worksheet,
            range: data.address,
            rowCount: data.rowCount,
            columnCount: data.columnCount,
            statistics: stats,
          });
        }
        break;
      }

      case "pivot_summary": {
        // Get pivot tables in the workbook
        const pivotEndpoint = `${workbookBase}/worksheets/${worksheet}/pivotTables`;
        const pivotResponse = await client.get<any>(pivotEndpoint);

        if (pivotResponse.success && pivotResponse.data) {
          const pivotTables = (pivotResponse.data as any).value || [];

          const summaries = await Promise.all(
            pivotTables.map(async (pivot: any) => {
              const refreshEndpoint = `${workbookBase}/worksheets/${worksheet}/pivotTables/${pivot.id}/refresh`;
              await client.post(refreshEndpoint, {});

              return {
                id: pivot.id,
                name: pivot.name,
                worksheet: pivot.worksheet?.name,
              };
            }),
          );

          return jsonTextResponse({
            analysisType: "pivot_summary",
            worksheet,
            pivotTableCount: pivotTables.length,
            pivotTables: summaries,
          });
        }
        break;
      }

      case "data_validation": {
        // Get data validation rules
        const validationEndpoint = range
          ? `${workbookBase}/worksheets/${worksheet}/range(address='${range}')/dataValidation`
          : `${workbookBase}/worksheets/${worksheet}/usedRange/dataValidation`;

        const validationResponse = await client.get<any>(validationEndpoint);

        if (validationResponse.success && validationResponse.data) {
          const validation = validationResponse.data;

          return jsonTextResponse({
            analysisType: "data_validation",
            worksheet,
            range: range || "usedRange",
            validation: {
              errorAlert: validation.errorAlert,
              errorMessage: validation.errorMessage,
              errorTitle: validation.errorTitle,
              operator: validation.operator,
              type: validation.type,
              formula1: validation.formula1,
              formula2: validation.formula2,
            },
          });
        }
        break;
      }

      case "named_ranges": {
        // Get all named ranges in the workbook
        const namedRangesEndpoint = `${workbookBase}/names`;
        const namedRangesResponse = await client.get<any>(namedRangesEndpoint);

        if (namedRangesResponse.success && namedRangesResponse.data) {
          const namedRanges = (namedRangesResponse.data as any).value || [];

          return jsonTextResponse({
            analysisType: "named_ranges",
            count: namedRanges.length,
            namedRanges: namedRanges.map((nr: any) => ({
              name: nr.name,
              type: nr.type,
              value: nr.value,
              visible: nr.visible,
              scope: nr.scope,
            })),
          });
        }
        break;
      }

      case "used_range": {
        // Get the used range of the worksheet
        const usedRangeEndpoint = `${workbookBase}/worksheets/${worksheet}/usedRange`;
        const usedRangeResponse = await client.get<any>(usedRangeEndpoint);

        if (usedRangeResponse.success && usedRangeResponse.data) {
          const usedRange = usedRangeResponse.data;

          // Calculate some basic info about the data
          const values = usedRange.values || [];
          let emptyCells = 0;
          let numberCells = 0;
          let textCells = 0;
          let formulaCells = 0;

          for (let row of values) {
            for (let cell of row) {
              if (cell === null || cell === "") {
                emptyCells++;
              } else if (typeof cell === "number") {
                numberCells++;
              } else if (typeof cell === "string") {
                textCells++;
              }
            }
          }

          // Check formulas
          const formulas = usedRange.formulas || [];
          for (let row of formulas) {
            for (let formula of row) {
              if (formula && formula.startsWith("=")) {
                formulaCells++;
              }
            }
          }

          return jsonTextResponse({
            analysisType: "used_range",
            worksheet,
            usedRange: {
              address: usedRange.address,
              rowCount: usedRange.rowCount,
              columnCount: usedRange.columnCount,
              totalCells: usedRange.rowCount * usedRange.columnCount,
              cellTypes: {
                empty: emptyCells,
                numbers: numberCells,
                text: textCells,
                formulas: formulaCells,
              },
            },
          });
        }
        break;
      }

      default:
        throw new Error(`Invalid analysis type: ${analysisType}`);
    }

    throw new Error(`Failed to perform analysis: ${analysisType}`);
  } catch (error) {
    return toolErrorResponse("excel_analysis", error);
  }
}

// Export all Excel tools and handlers
export const excelTools = [excelOperations, excelAnalysis];

export const excelHandlers = {
  excel_operations: handleExcelOperations,
  excel_analysis: handleExcelAnalysis,
};
