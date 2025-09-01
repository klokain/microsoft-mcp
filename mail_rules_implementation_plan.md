# Exchange Online Mail Rules Implementation Plan

## Overview
Implement comprehensive mail rules management for Exchange Online using Microsoft Graph API's messageRule endpoints. This will enable users to create, manage, and execute inbox rules programmatically.

## API Endpoints
- `GET /me/mailFolders/inbox/messageRules` - List all rules
- `GET /me/mailFolders/inbox/messageRules/{id}` - Get specific rule
- `POST /me/mailFolders/inbox/messageRules` - Create new rule
- `PATCH /me/mailFolders/inbox/messageRules/{id}` - Update existing rule
- `DELETE /me/mailFolders/inbox/messageRules/{id}` - Delete rule

## Implementation Steps

### 1. Add Helper Functions to `graph.py`

```python
def get_mail_folder_id(folder_name: str, account_id: str | None = None) -> str:
    """Get folder ID by name or return if already an ID"""
    if folder_name.startswith("AAMkA"):  # Already an ID
        return folder_name
    
    # Try well-known folders first
    well_known = ["inbox", "drafts", "sentitems", "deleteditems", "archive", "junkemail"]
    if folder_name.lower() in well_known:
        return folder_name
    
    # Search for folder by display name
    folders = request_paginated("/me/mailFolders", account_id)
    for folder in folders:
        if folder["displayName"].lower() == folder_name.lower():
            return folder["id"]
    
    raise ValueError(f"Folder '{folder_name}' not found")

def validate_rule_conditions(conditions: dict[str, Any]) -> dict[str, Any]:
    """Validate and format rule conditions"""
    valid_conditions = {
        "senderContains", "recipientContains", "subjectContains", 
        "subjectOrBodyContains", "bodyContains", "headerContains",
        "fromAddresses", "sentToAddresses", "categories",
        "importance", "sensitivity", "withinSizeRange",
        "receivedDateTimeRange", "isApprovalRequest", "isAutomaticForward",
        "isAutomaticReply", "isEncrypted", "isMeetingRequest",
        "isMeetingResponse", "isNonDeliveryReport", "isPermissionControlled",
        "isReadReceipt", "isSigned", "isVoicemail", "hasAttachments"
    }
    
    validated = {}
    for key, value in conditions.items():
        if key in valid_conditions:
            validated[key] = value
    
    return validated

def validate_rule_actions(actions: dict[str, Any], account_id: str | None = None) -> dict[str, Any]:
    """Validate and format rule actions"""
    valid_actions = {
        "moveToFolder", "copyToFolder", "delete", "permanentDelete",
        "markAsRead", "markImportance", "forwardTo", "forwardAsAttachmentTo",
        "redirectTo", "assignCategories", "stopProcessingRules"
    }
    
    validated = {}
    for key, value in actions.items():
        if key not in valid_actions:
            continue
            
        # Convert folder names to IDs for folder actions
        if key in ["moveToFolder", "copyToFolder"] and value:
            validated[key] = get_mail_folder_id(value, account_id)
        # Ensure email addresses are in correct format for forward/redirect
        elif key in ["forwardTo", "forwardAsAttachmentTo", "redirectTo"]:
            if isinstance(value, str):
                validated[key] = [{"emailAddress": {"address": value}}]
            elif isinstance(value, list):
                validated[key] = [{"emailAddress": {"address": email}} for email in value]
        else:
            validated[key] = value
    
    return validated
```

### 2. Add Mail Rule Tools to `tools.py`

```python
@mcp.tool
def list_mail_rules(
    account_id: str,
    include_disabled: bool = True
) -> list[dict[str, Any]]:
    """
    List all mail rules for the user's inbox
    
    Args:
        account_id: Microsoft account ID
        include_disabled: Whether to include disabled rules (default: True)
    
    Returns:
        List of mail rules with their configurations
    """
    rules = list(request_paginated("/me/mailFolders/inbox/messageRules", account_id))
    
    if not include_disabled:
        rules = [r for r in rules if r.get("isEnabled", False)]
    
    return rules

@mcp.tool
def get_mail_rule(
    account_id: str,
    rule_id: str
) -> dict[str, Any]:
    """
    Get details of a specific mail rule
    
    Args:
        account_id: Microsoft account ID
        rule_id: The ID of the mail rule
    
    Returns:
        Mail rule details including conditions and actions
    """
    result = request("GET", f"/me/mailFolders/inbox/messageRules/{rule_id}", account_id)
    if not result:
        raise ValueError(f"Mail rule {rule_id} not found")
    return result

@mcp.tool
def create_mail_rule(
    account_id: str,
    display_name: str,
    conditions: dict[str, Any],
    actions: dict[str, Any],
    is_enabled: bool = True,
    sequence: int | None = None
) -> dict[str, Any]:
    """
    Create a new mail rule
    
    Args:
        account_id: Microsoft account ID
        display_name: Name for the rule
        conditions: Rule conditions (e.g., {"senderContains": ["newsletter"], "hasAttachments": True})
        actions: Rule actions (e.g., {"moveToFolder": "Archive", "markAsRead": True})
        is_enabled: Whether rule is active (default: True)
        sequence: Rule priority/order (lower numbers run first)
    
    Supported conditions:
        - senderContains, recipientContains, subjectContains, bodyContains
        - subjectOrBodyContains, fromAddresses, sentToAddresses
        - hasAttachments, importance, sensitivity, categories
        - isApprovalRequest, isMeetingRequest, isAutomaticForward, etc.
    
    Supported actions:
        - moveToFolder, copyToFolder, delete, permanentDelete
        - markAsRead, markImportance, forwardTo, redirectTo
        - assignCategories, stopProcessingRules
    
    Returns:
        Created mail rule object
    """
    validated_conditions = validate_rule_conditions(conditions)
    validated_actions = validate_rule_actions(actions, account_id)
    
    if not validated_conditions:
        raise ValueError("At least one valid condition is required")
    if not validated_actions:
        raise ValueError("At least one valid action is required")
    
    payload = {
        "displayName": display_name,
        "conditions": validated_conditions,
        "actions": validated_actions,
        "isEnabled": is_enabled
    }
    
    if sequence is not None:
        payload["sequence"] = sequence
    
    result = request("POST", "/me/mailFolders/inbox/messageRules", account_id, json=payload)
    if not result:
        raise ValueError("Failed to create mail rule")
    
    return result

@mcp.tool
def update_mail_rule(
    account_id: str,
    rule_id: str,
    display_name: str | None = None,
    conditions: dict[str, Any] | None = None,
    actions: dict[str, Any] | None = None,
    is_enabled: bool | None = None,
    sequence: int | None = None
) -> dict[str, Any]:
    """
    Update an existing mail rule
    
    Args:
        account_id: Microsoft account ID
        rule_id: The ID of the mail rule to update
        display_name: New name for the rule (optional)
        conditions: New conditions (optional, replaces all conditions)
        actions: New actions (optional, replaces all actions)
        is_enabled: Enable/disable the rule (optional)
        sequence: New priority order (optional)
    
    Returns:
        Updated mail rule object
    """
    payload = {}
    
    if display_name is not None:
        payload["displayName"] = display_name
    
    if conditions is not None:
        payload["conditions"] = validate_rule_conditions(conditions)
    
    if actions is not None:
        payload["actions"] = validate_rule_actions(actions, account_id)
    
    if is_enabled is not None:
        payload["isEnabled"] = is_enabled
    
    if sequence is not None:
        payload["sequence"] = sequence
    
    if not payload:
        raise ValueError("No updates provided")
    
    result = request("PATCH", f"/me/mailFolders/inbox/messageRules/{rule_id}", 
                    account_id, json=payload)
    if not result:
        raise ValueError(f"Failed to update mail rule {rule_id}")
    
    return result

@mcp.tool
def delete_mail_rule(
    account_id: str,
    rule_id: str
) -> dict[str, Any]:
    """
    Delete a mail rule
    
    Args:
        account_id: Microsoft account ID
        rule_id: The ID of the mail rule to delete
    
    Returns:
        Success status
    """
    request("DELETE", f"/me/mailFolders/inbox/messageRules/{rule_id}", account_id)
    return {"success": True, "message": f"Mail rule {rule_id} deleted successfully"}

@mcp.tool
def toggle_mail_rule(
    account_id: str,
    rule_id: str,
    enabled: bool | None = None
) -> dict[str, Any]:
    """
    Enable or disable a mail rule
    
    Args:
        account_id: Microsoft account ID
        rule_id: The ID of the mail rule
        enabled: True to enable, False to disable, None to toggle
    
    Returns:
        Updated rule status
    """
    if enabled is None:
        # Get current status and toggle
        rule = get_mail_rule(account_id, rule_id)
        enabled = not rule.get("isEnabled", False)
    
    result = request("PATCH", f"/me/mailFolders/inbox/messageRules/{rule_id}", 
                    account_id, json={"isEnabled": enabled})
    
    return {
        "success": True,
        "rule_id": rule_id,
        "is_enabled": enabled,
        "message": f"Rule {'enabled' if enabled else 'disabled'} successfully"
    }
```

### 3. Add Tests to `test_integration.py`

```python
def test_mail_rules_crud():
    """Test creating, reading, updating, and deleting mail rules"""
    account_id = get_test_account_id()
    
    # Create a test rule
    rule = create_mail_rule(
        account_id,
        display_name="Test Newsletter Rule",
        conditions={
            "senderContains": ["newsletter", "noreply"],
            "subjectContains": ["unsubscribe"]
        },
        actions={
            "moveToFolder": "Junk Email",
            "markAsRead": True
        },
        is_enabled=False  # Don't enable during testing
    )
    
    assert rule["displayName"] == "Test Newsletter Rule"
    assert not rule["isEnabled"]
    rule_id = rule["id"]
    
    try:
        # List rules
        rules = list_mail_rules(account_id)
        assert any(r["id"] == rule_id for r in rules)
        
        # Get specific rule
        fetched_rule = get_mail_rule(account_id, rule_id)
        assert fetched_rule["id"] == rule_id
        assert "newsletter" in fetched_rule["conditions"]["senderContains"]
        
        # Update rule
        updated = update_mail_rule(
            account_id,
            rule_id,
            display_name="Updated Test Rule",
            is_enabled=True
        )
        assert updated["displayName"] == "Updated Test Rule"
        assert updated["isEnabled"]
        
        # Toggle rule
        toggled = toggle_mail_rule(account_id, rule_id, enabled=False)
        assert not toggled["is_enabled"]
        
    finally:
        # Clean up
        delete_mail_rule(account_id, rule_id)
        
        # Verify deletion
        rules = list_mail_rules(account_id)
        assert not any(r["id"] == rule_id for r in rules)

def test_complex_mail_rule():
    """Test creating a complex mail rule with multiple conditions and actions"""
    account_id = get_test_account_id()
    
    rule = create_mail_rule(
        account_id,
        display_name="Complex Priority Rule",
        conditions={
            "fromAddresses": [
                {"emailAddress": {"address": "boss@company.com"}},
                {"emailAddress": {"address": "ceo@company.com"}}
            ],
            "importance": "high",
            "hasAttachments": True,
            "subjectOrBodyContains": ["urgent", "ASAP", "priority"]
        },
        actions={
            "markImportance": "high",
            "assignCategories": ["Important", "Review"],
            "forwardTo": "assistant@company.com",
            "stopProcessingRules": True
        },
        sequence=1  # High priority
    )
    
    try:
        assert rule["sequence"] == 1
        assert rule["actions"]["markImportance"] == "high"
        assert "Important" in rule["actions"]["assignCategories"]
        
    finally:
        delete_mail_rule(account_id, rule["id"])
```

### 4. Update Documentation

#### README.md
Add new section after Calendar Tools:

```markdown
### Mail Rule Tools
- **`list_mail_rules`** - List all inbox rules
- **`get_mail_rule`** - Get specific rule details
- **`create_mail_rule`** - Create new mail rule with conditions and actions
- **`update_mail_rule`** - Modify existing mail rules
- **`delete_mail_rule`** - Remove mail rules
- **`toggle_mail_rule`** - Enable/disable rules
```

Add usage examples:

```markdown
# Mail rules examples
> create a rule to move all newsletters to a folder
> list all my active mail rules
> disable the rule that forwards to my assistant
> create a rule to flag emails from my boss as important
```

Add code example:

```python
### Automated Email Organization
# Create rule to organize newsletters
create_mail_rule(
    account_id,
    "Newsletter Management",
    conditions={
        "senderContains": ["newsletter", "noreply", "marketing"],
        "subjectContains": ["unsubscribe", "weekly", "digest"]
    },
    actions={
        "moveToFolder": "Newsletters",
        "markAsRead": False,
        "assignCategories": ["Newsletters"]
    }
)

# Priority email handling
create_mail_rule(
    account_id,
    "VIP Emails",
    conditions={
        "fromAddresses": [
            {"emailAddress": {"address": "ceo@company.com"}},
            {"emailAddress": {"address": "important.client@example.com"}}
        ],
        "importance": "high"
    },
    actions={
        "markImportance": "high",
        "assignCategories": ["VIP", "Urgent"],
        "forwardTo": "assistant@company.com"
    },
    sequence=1  # Process first
)
```

#### CLAUDE.md
Add section on mail rules:

```markdown
## Mail Rules Management

The server supports comprehensive mail rule management through the Microsoft Graph messageRules API:

### Rule Conditions
- Text matching: senderContains, subjectContains, bodyContains
- Address matching: fromAddresses, sentToAddresses
- Properties: hasAttachments, importance, sensitivity
- Message types: isMeetingRequest, isAutomaticForward

### Rule Actions
- Folder operations: moveToFolder, copyToFolder
- Message operations: delete, markAsRead, markImportance
- Forwarding: forwardTo, redirectTo, forwardAsAttachmentTo
- Organization: assignCategories
- Processing: stopProcessingRules

### Implementation Notes
- Rules are processed in sequence order (lower numbers first)
- Folder names are automatically converted to IDs
- Email addresses are formatted to Graph API requirements
- Rules can be enabled/disabled without deletion
```

### 5. Error Handling Enhancements

Add to `graph.py`:

```python
class MailRuleError(Exception):
    """Custom exception for mail rule operations"""
    pass

def handle_rule_error(response: httpx.Response) -> None:
    """Handle mail rule specific errors"""
    if response.status_code == 400:
        error = response.json().get("error", {})
        message = error.get("message", "Bad request")
        
        if "duplicate" in message.lower():
            raise MailRuleError(f"A rule with this name already exists")
        elif "condition" in message.lower():
            raise MailRuleError(f"Invalid rule condition: {message}")
        elif "action" in message.lower():
            raise MailRuleError(f"Invalid rule action: {message}")
        else:
            raise MailRuleError(message)
```

## Testing Strategy

1. **Unit Tests**: Test validation functions for conditions and actions
2. **Integration Tests**: Test CRUD operations with real API
3. **Edge Cases**: 
   - Maximum number of rules (256 per mailbox)
   - Rule priority conflicts
   - Invalid folder names
   - Circular forwarding detection

## Usage Examples for AI Assistants

```python
# Example 1: Auto-archive old emails
create_mail_rule(
    account_id,
    "Archive Old Emails",
    conditions={
        "receivedDateTimeRange": {
            "startDateTime": "2024-01-01T00:00:00Z",
            "endDateTime": "2024-06-01T00:00:00Z"
        }
    },
    actions={
        "moveToFolder": "Archive"
    }
)

# Example 2: Spam filtering
create_mail_rule(
    account_id,
    "Spam Filter",
    conditions={
        "subjectOrBodyContains": ["win", "free", "click here", "limited offer"],
        "senderContains": ["@suspicious-domain.com"]
    },
    actions={
        "moveToFolder": "Junk Email",
        "permanentDelete": False
    }
)

# Example 3: Meeting request handling
create_mail_rule(
    account_id,
    "Meeting Requests",
    conditions={
        "isMeetingRequest": True,
        "importance": "high"
    },
    actions={
        "assignCategories": ["Meetings", "Calendar"],
        "markImportance": "high"
    }
)
```

## Benefits

1. **Automation**: Reduce manual email management
2. **Consistency**: Apply rules uniformly across all emails
3. **Efficiency**: Process emails immediately on arrival
4. **Organization**: Automatic categorization and filing
5. **Prioritization**: Highlight important emails automatically

## Limitations

- Maximum 256 rules per mailbox
- Rules only apply to inbox (not other folders)
- Some actions require specific Exchange Online licenses
- Rules execute in order, can impact performance with many rules