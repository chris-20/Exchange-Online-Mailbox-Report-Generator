# Exchange Online Mailbox Report Generator

## Description
A PowerShell script that generates comprehensive reports of Exchange Online mailboxes. The script features smart module management, automatic connection handling, and detailed progress tracking. It's designed for Exchange Online administrators who need to efficiently collect and analyze mailbox statistics.

## Features
- 🔄 Smart module management
  - Auto-detection of Exchange Online Management module
  - Optional installation if missing
  - Update checks with optional module updates
- 🔌 Intelligent connection handling
  - Detects existing Exchange Online sessions
  - Prevents unnecessary authentication prompts
  - Option to maintain connection for subsequent tasks
- 📊 Comprehensive mailbox reporting
  - Display name and email address
  - Mailbox type and status
  - Total item size and count
  - Last logon time
  - Database location
  - Enabled status
- 📈 Progress tracking
  - Real-time processing status
  - Detailed progress bar
  - Color-coded status messages
- 📁 Automated export
  - CSV export with timestamp
  - Export to script execution directory
  - Full path confirmation

## Requirements
- Windows PowerShell 5.1 or PowerShell 7+
- Exchange Online Management module (auto-install available)
- Exchange Online administration credentials
- Internet connectivity

## Installation
1. Download the script
2. Ensure you have appropriate Exchange Online permissions
3. Run the script in PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File .\ExchangeOnlineMailboxReport.ps1
```

## Usage
The script will:
1. Check for and optionally install required modules
2. Verify/establish Exchange Online connection
3. Collect mailbox data
4. Display results in console
5. Export to CSV
6. Optionally maintain the connection

## Output Example
The CSV output includes:
```
DisplayName,EmailAddress,MailboxType,TotalItemSize,ItemCount,LastLogonTime,Database,Enabled
John Doe,john.doe@company.com,UserMailbox,5.5 GB (5,911,495,680 bytes),10521,2024-03-19 09:45:12,DB01,True
```

## Best Practices
- Run during off-peak hours for large organizations
- Review module updates before accepting
- Keep existing connections if running multiple Exchange tasks
- Save reports with meaningful timestamps

## Error Handling
- Module installation verification
- Connection state validation
- Mailbox access permissions check
- Export path verification
- Detailed error messages

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

## License
MIT License - feel free to use and modify for your needs.

## Version History
- 1.0.0 (2024-03-19)
  - Initial release
  - Basic reporting functionality
  - Smart connection handling
  - Module management

## Acknowledgments
- Microsoft Exchange Online documentation
- PowerShell community best practices
- Exchange Online Management module team

## Support
For issues or questions:
1. Open an issue in this repository
2. Provide error messages and script version
3. Describe your Exchange Online environment
