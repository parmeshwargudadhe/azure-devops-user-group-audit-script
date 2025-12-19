# Azure DevOps User Group Audit Tool

A Python script to audit Azure DevOps user group memberships and export results to Excel format.

python -m venv venv

👉 PowerShell Execution Policy Error: use following commmand
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

.\venv\Scripts\Activate.ps1

pip install -r requirements.txt

## Features

- 🔍 **Search specific users** - Look up group memberships for individual users
- 📊 **Full organization audit** - Process all users (900+ supported) with progress tracking
- 💾 **Excel exports** - Save results to formatted Excel files with auto-adjusted columns
- 📋 **Quick user count** - Get total number of users in the organization
- ⚡ **Optimized performance** - Built-in rate limiting and error handling

## Quick Start

1. **Install dependencies:**
   ```bash
   pip install requests openpyxl
   ```

2. **Configure the script:**
   ```python
   # In audit_tool.py, update:
   ORG = "your-organization-name"
   PAT = "your-personal-access-token"  # Needs 'User Profile Read' scope
   ```

3. **Run the tool:**
   ```bash
   python audit_tool.py
   ```

## Usage

When you run the script, you'll see a menu with 4 options:

```
1. 🔍 Search for a specific user's group memberships
2. 📊 Run full audit for all users
3. 📋 Get total user count
4. ❌ Exit
```

### Option 1: Search Specific User
- Enter user email address
- View group memberships with scope details
- Option to save results to Excel

### Option 2: Full Audit
- Processes all users in the organization
- Shows real-time progress
- Automatically exports to timestamped Excel file
- Option to open the file after completion

### Option 3: User Count
- Quick count of all users in the organization

### Option 4: Exit
- Close the application

## Output Format

Excel files include:
- **User Email** - User's email address
- **User Display Name** - User's display name in Azure DevOps
- **Group Name** - Name of the group
- **Group Principal Name** - Technical group identifier
- **Scope Type** - Organization or Project level
- **Scope Name** - Organization or project name
- **Group Description** - Group description (if available)

Files are automatically timestamped: `ado_groups_username_20241219_143022.xlsx`

## Security Notes

⚠️ **Important Security Considerations:**
- Never commit your PAT token to version control
- The PAT only needs `User Profile Read` scope (minimum permissions)
- Consider using environment variables or separate config files for credentials
- Excel output files contain user information - handle appropriately

## Requirements

- Python 3.6+
- Azure DevOps organization
- PAT with `User Profile Read` scope
- Internet connection to Azure DevOps APIs

## API Endpoints Used

- `GET /_apis/graph/users` - List all users
- `GET /_apis/graph/memberships/{user}` - Get user's group memberships
- `GET /_apis/graph/groups/{group}` - Get group details

## Troubleshooting

**Authentication errors:** Verify your PAT has correct scope and hasn't expired.

**Rate limiting:** The script includes built-in delays to avoid API limits.

**No groups found:** Some users may not have explicit group memberships.

**Large organizations:** Full audit for 900+ users takes approximately 15-30 minutes.

## License

MIT License - See LICENSE file for details.
