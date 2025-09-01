# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Microsoft Graph MCP (Model Context Protocol) server that provides AI assistants access to Microsoft 365 services including Outlook, Calendar, OneDrive, and Contacts. It's a fork maintained by resulta.tech based on the original work by elyxlz.

The server implements multi-account support, allowing users to authenticate and manage multiple Microsoft accounts (personal, work, school) simultaneously through a single MCP interface.

## Development Commands

```bash
# Install dependencies
uv sync

# Run the MCP server locally
uv run microsoft-mcp

# Run tests
uv run pytest tests/ -v

# Type checking
uv run pyright

# Format code
uvx ruff format .

# Lint and fix issues
uvx ruff check --fix --unsafe-fixes .

# Authentication setup for development
export MICROSOFT_MCP_CLIENT_ID="your-app-id-here"
uv run authenticate.py
```

## Architecture Overview

### Core Components

- **`server.py`** - Entry point that validates environment variables and starts the FastMCP server
- **`tools.py`** - Contains all MCP tool definitions using FastMCP decorators, organized by service area (email, calendar, contacts, files)
- **`auth.py`** - Handles Microsoft authentication using MSAL (Microsoft Authentication Library) with token caching
- **`graph.py`** - Low-level Microsoft Graph API client with retry logic, pagination, and error handling

### Key Architectural Patterns

1. **Multi-Account Architecture**: All tools require an `account_id` parameter as the first argument to specify which authenticated Microsoft account to use
2. **Token Management**: Uses MSAL SerializableTokenCache stored in `~/.microsoft_mcp_token_cache.json` for persistent authentication
3. **Pagination Support**: Graph API responses are handled with automatic pagination through `request_paginated()`
4. **Retry Logic**: Implements exponential backoff for rate limiting (429) and server errors (5xx)
5. **Chunked Uploads**: Large file uploads to OneDrive use chunked upload sessions (15 x 320 KiB chunks)
6. **Batch Operations**: Email batch operations using Microsoft Graph JSON batching with automatic concurrency handling (max 4 concurrent requests per mailbox)

### Service Areas

The codebase is organized around four main Microsoft Graph service areas:

- **Email**: List, send, reply, manage attachments, folder operations, batch operations (delete, move, update multiple emails)
- **Calendar**: Event creation/management, availability checking, invitation responses
- **Contacts**: CRUD operations for contact management
- **Files (OneDrive)**: File/folder operations with upload/download support

## Environment Variables

Required:
- `MICROSOFT_MCP_CLIENT_ID` - Azure application ID (required for startup)

Optional:
- `MICROSOFT_MCP_TENANT_ID` - Defaults to "common" (use "consumers" for personal accounts only)

## Authentication Flow

1. User calls `authenticate_account()` to get device code and verification URL
2. User visits URL and enters device code in browser
3. User calls `complete_authentication(flow_cache)` to finalize authentication
4. Token is cached locally for subsequent requests
5. All subsequent tool calls use `account_id` to specify which account to use

## Testing

- Integration tests in `tests/test_integration.py` use the MCP stdio client to test against the actual server
- Tests require a valid `MICROSOFT_MCP_CLIENT_ID` environment variable
- Helper function `parse_result()` normalizes FastMCP response formats
- Batch operation tests validate proper error handling and cleanup

## Batch Operations

The server includes efficient batch operations for bulk email management:

### Available Batch Tools
- **`batch_delete_emails()`** - Delete multiple emails at once
- **`batch_move_emails()`** - Move multiple emails to a folder
- **`batch_update_emails()`** - Update multiple emails (mark read/unread, etc.)

### Implementation Details
- Uses Microsoft Graph `/$batch` endpoint for optimal performance
- Automatically handles mailbox concurrency limits (max 4 concurrent per mailbox)
- Uses `dependsOn` properties to ensure sequential processing when needed
- Returns detailed results with success/failure status for each operation
- Supports up to 20 operations per batch (Graph API limit)

### Concurrency Management
- Splits large batches into chunks of 4 to respect Outlook's mailbox concurrency limit
- Uses dependency chaining to avoid "TooManyRequests" errors
- Implements proper error handling and retry logic with exponential backoff

## Microsoft Graph API Integration

- Base URL: `https://graph.microsoft.com/v1.0`
- Uses standard OAuth 2.0 device flow for authentication
- Implements Microsoft Graph API best practices:
  - ConsistencyLevel headers for search queries
  - Prefer headers for email body content type
  - Proper retry handling for throttling
  - Chunked upload for large files
  - JSON batching for bulk operations with concurrency management

## Development Notes

- Built on FastMCP framework for MCP server implementation
- Uses httpx for HTTP client with 30-second timeout
- Token cache file location: `~/.microsoft_mcp_token_cache.json`
- Upload chunk size configured for optimal OneDrive performance
- All datetime handling uses ISO format strings