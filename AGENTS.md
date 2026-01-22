# Agent 365 MCP - Agent Integration Guide

This document provides guidance for AI agents integrating with Microsoft 365 via the Agent 365 MCP server.

## Quick Reference

### Available Tool Prefixes

| Prefix | Service | Best For |
|--------|---------|----------|
| `sharepoint_*` | SharePoint/OneDrive | Finding and listing files |
| `word_*` | Word | Reading/writing .docx documents |
| `excel_*` | Excel | Reading/writing .xlsx spreadsheets |
| `teams_*` | Teams | Chats, channels, meeting info |
| `mail_*` | Outlook | Email read/send/search |
| `calendar_*` | Calendar | Events and scheduling |
| `me_*` | Profile | User and org info |
| `copilot_*` | M365 Copilot | Semantic search, summaries |

## Tool Selection Guide

### When Given a SharePoint/OneDrive URL

```
URL ends with .docx → Use word_WordGetDocumentContent
URL ends with .xlsx → Use excel_ExcelGetDocumentContent
URL ends with .pptx → Use sharepoint_readSmallTextFile (limited support)
Other files → Use sharepoint_readSmallTextFile
Unknown → Use sharepoint_getFileOrFolderMetadataByUrl first
```

### When Searching for Files

1. **Know the filename**: `sharepoint_findFileOrFolder`
2. **Don't know filename**: `copilot_CopilotChat` with search query
3. **Need recent files**: `sharepoint_getFolderChildren` on known folder

### When Reading Documents

```
Word Document (.docx):
  - Prefer: word_WordGetDocumentContent
  - Returns: Full text, comments, structure
  - Avoid: sharepoint_readSmallTextFile (returns raw XML)

Excel Spreadsheet (.xlsx):
  - Prefer: excel_ExcelGetDocumentContent
  - Returns: Cell data as structured text
  - Avoid: sharepoint_readSmallTextFile (returns binary)

Other Text Files:
  - Use: sharepoint_readSmallTextFile
```

### When Searching for Information

1. **Know where to look**: Use specific service tools
2. **Don't know where**: `copilot_CopilotChat` - searches across all M365

```
copilot_CopilotChat capabilities:
- Search emails, files, chats simultaneously
- Summarize meeting transcripts
- Find information by semantic meaning
- Cite sources with links
```

### When Working with Teams

```
Find a chat:
  teams_listChats (with expand=lastMessagePreview for recent activity)

Read messages:
  teams_listChatMessages (needs chat-id from listChats)

Find meeting chat:
  teams_listChats → filter by chatType="meeting" and topic

Post message (requires user confirmation):
  teams_postMessage / teams_postChannelMessage
```

### When Working with Email

```
Search emails:
  mail_SearchMessagesAsync (more efficient than listing all)

Read specific email:
  mail_GetMessageAsync (needs message ID)

Send email (requires user confirmation):
  mail_SendEmailWithAttachmentsAsync
```

### When Working with Calendar

```
List events:
  calendar_listCalendarView (use date range to limit results)

Create event (requires user confirmation):
  calendar_createEvent
```

## Handling Large Responses

Some tools return large responses. The proxy handles this automatically:

1. **If AGENT365_LARGE_FILE_DIR is set**: Large responses are saved to files
   - Agent receives filepath in response
   - Use the Read tool to access the file

2. **If not set**: Responses are truncated
   - Agent receives partial content + hint
   - Recommend: Use more specific queries

### Tips to Reduce Response Size

```
Calendar:
  - Use specific date ranges instead of listing all events
  - Use top parameter to limit results

Teams:
  - Use filter parameter when searching chats
  - Paginate with top parameter

SharePoint:
  - Search by filename instead of listing folders
  - Use Copilot for broad searches
```

## Common Patterns

### Find and Read a Document

```javascript
// 1. Find the file
sharepoint_findFileOrFolder({ searchQuery: "quarterly report" })

// 2. Check file type from result
// If .docx:
word_WordGetDocumentContent({ url: "sharepoint_url_from_step_1" })

// If .xlsx:
excel_ExcelGetDocumentContent({ url: "sharepoint_url_from_step_1" })
```

### Get Meeting Notes

```javascript
// 1. Search with Copilot for meeting info
copilot_CopilotChat({
  message: "What was discussed in the team meeting on January 15?"
})

// Or manually:
// 2. Find meeting chat
teams_listChats({
  filter: "chatType eq 'meeting'",
  expand: "lastMessagePreview"
})

// 3. Get messages from that chat
teams_listChatMessages({
  "chat-id": "chat_id_from_step_2"
})
```

### Send Email with Context

```javascript
// 1. Get user's profile for proper greeting
me_getMyProfile()

// 2. Search for relevant context
copilot_CopilotChat({
  message: "Find recent emails about project X"
})

// 3. Compose and send (requires user confirmation)
mail_SendEmailWithAttachmentsAsync({
  // ... email details
})
```

## Error Handling

### "Server is disabled"

Some servers may be disabled due to:
- Missing Copilot license
- Permission denied
- Server unavailable

Check `AGENT365_DISABLED_SERVERS` configuration or server logs.

### "Response too large"

Use more specific queries or enable `AGENT365_LARGE_FILE_DIR`.

### "Token expired"

User needs to re-authenticate:
```bash
npx github:rapyuta-robotics/agent365-mcp auth
```

## Skills Integration

### Relevant Skills

When working with Agent 365, consider using these skills:

- **brainstorming**: Before complex M365 tasks, brainstorm approach
- **systematic-debugging**: When M365 operations fail unexpectedly
- **writing-plans**: For multi-step M365 workflows

### Example: Document Creation Workflow

```
1. Use brainstorming skill to clarify requirements
2. Search existing content with copilot_CopilotChat
3. Create document with word_WordCreateNewDocument
4. Add to SharePoint with sharepoint_createSmallTextFile
5. Share with user confirmation
```

## Best Practices

1. **Check file type before reading**: Use appropriate tool for file format
2. **Use Copilot for discovery**: When you don't know exact locations
3. **Limit response size**: Use date ranges, filters, pagination
4. **Confirm before sending**: Always confirm emails/messages with user
5. **Handle disabled servers**: Gracefully fall back to alternatives
6. **Cite sources**: When using Copilot results, include attribution

## API Limits

- Request timeout: 60 seconds (configurable)
- Max response buffer: 10MB
- Token expiration: 1 hour
- No rate limiting from proxy (M365 may rate limit)

## Debugging

### Check Available Tools

```bash
# List all loaded tools
npx github:rapyuta-robotics/agent365-mcp status
```

### Check Server Logs

When running via an MCP client, server logs appear in stderr.

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| No tools loaded | Token expired | Re-authenticate |
| Copilot disabled | No license | Use other tools |
| Large response truncated | Response > threshold | Set LARGE_FILE_DIR |
| Permission denied | Missing consent | Admin grant consent |
