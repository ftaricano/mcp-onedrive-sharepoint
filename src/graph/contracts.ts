import { GraphResponse } from "./models.js";
import { GraphApiError, createUserFriendlyError } from "./error-handler.js";

export interface ToolErrorPayload {
  summary: string;
  error: {
    code: string;
    category: string;
    message: string;
    retryable: boolean;
    severity: string;
    statusCode?: number;
    context?: string;
    suggestedAction?: string;
  };
}

export interface PaginationInfo {
  limit?: number;
  returned: number;
  totalCount?: number;
  nextPageToken?: string;
  hasMore: boolean;
}

export interface PaginatedResult<T> {
  items: T[];
  pagination: PaginationInfo;
}

export function assertGraphPayloadHasNoError<T>(
  payload: T,
  context?: string,
): T {
  if (
    payload &&
    typeof payload === "object" &&
    "error" in (payload as Record<string, unknown>) &&
    (payload as { error?: unknown }).error
  ) {
    throw new GraphApiError(payload, context);
  }

  return payload;
}

export function extractPaginatedResult<T>(
  payload: GraphResponse<T> | { value?: T[] },
  limit?: number,
): PaginatedResult<T> {
  const response = assertGraphPayloadHasNoError(payload) as GraphResponse<T>;
  const items = Array.isArray(response?.value) ? response.value : [];
  const nextPageToken = response?.["@odata.nextLink"];

  return {
    items,
    pagination: {
      limit,
      returned: items.length,
      totalCount: response?.["@odata.count"],
      nextPageToken,
      hasMore: Boolean(nextPageToken),
    },
  };
}

export function jsonTextResponse(payload: unknown): {
  content: Array<{ type: "text"; text: string }>;
} {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(payload, null, 2),
      },
    ],
  };
}

export function jsonTextErrorResponse(payload: unknown): {
  content: Array<{ type: "text"; text: string }>;
  isError: true;
} {
  return {
    ...jsonTextResponse(payload),
    isError: true,
  };
}

export function toolErrorResponse(
  toolName: string,
  error: unknown,
  context?: string,
): {
  content: Array<{ type: "text"; text: string }>;
  isError: true;
} {
  const suffix = context ? ` (${context})` : "";
  const friendly = createUserFriendlyError(error, context);
  const summary = `Error in ${toolName}${suffix}: ${friendly}`;

  const graphError =
    error instanceof GraphApiError ? error : new GraphApiError(error, context);

  const payload: ToolErrorPayload = {
    summary,
    error: {
      code: graphError.code,
      category: graphError.category,
      message: graphError.message,
      retryable: graphError.isRetryable,
      severity: graphError.severity,
      statusCode: graphError.statusCode,
      context: graphError.context ?? context,
      suggestedAction: graphError.suggestedAction,
    },
  };

  return {
    content: [
      { type: "text", text: summary },
      { type: "text", text: JSON.stringify(payload, null, 2) },
    ],
    isError: true,
  };
}
