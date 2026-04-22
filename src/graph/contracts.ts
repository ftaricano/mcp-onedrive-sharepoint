import { GraphResponse } from "./models.js";
import { GraphApiError, createUserFriendlyError } from "./error-handler.js";

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

  return {
    content: [
      {
        type: "text",
        text: `Error in ${toolName}${suffix}: ${createUserFriendlyError(error, context)}`,
      },
    ],
    isError: true,
  };
}
