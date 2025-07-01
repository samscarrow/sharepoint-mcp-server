# SharePoint MCP Server

A Model Context Protocol server for browsing and interacting with Microsoft SharePoint sites and documents.

This TypeScript-based MCP server provides comprehensive SharePoint integration through Microsoft Graph API, enabling:

- **Resources**: Access SharePoint sites as MCP resources with structured metadata
- **Tools**: Search files, list sites, browse document libraries, and retrieve file content
- **Authentication**: Secure OAuth2 client credentials flow with automatic token management

## Features

### Resources
- List SharePoint sites accessible to the application
- Access individual site information via `sharepoint://site/{siteId}` URIs
- JSON-formatted site metadata with display names and web URLs

### Tools

#### `search_files`
Search for files and documents across SharePoint using Microsoft Graph Search API
- **Parameters**: 
  - `query` (required): Search query string
  - `limit` (optional): Maximum results to return (default: 10)

#### `list_sites`
List SharePoint sites accessible to the application
- **Parameters**:
  - `search` (optional): Filter sites by display name

#### `get_site_info`
Get detailed information about a specific SharePoint site
- **Parameters**:
  - `siteUrl` (required): SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)

#### `list_site_drives`
List document libraries (drives) in a SharePoint site
- **Parameters**:
  - `siteUrl` (required): SharePoint site URL

#### `list_drive_items`
List files and folders in a SharePoint document library
- **Parameters**:
  - `siteUrl` (required): SharePoint site URL
  - `driveId` (optional): Specific drive ID (uses default drive if not specified)
  - `folderPath` (optional): Folder path to list items from (default: root)

#### `get_file_content`
Get the content of a specific file from SharePoint (text files only)
- **Parameters**:
  - `siteUrl` (required): SharePoint site URL
  - `filePath` (required): Path to the file
  - `driveId` (optional): Specific drive ID (uses default drive if not specified)

## Prerequisites

### Azure App Registration
1. Register an application in Azure Active Directory
2. Configure API permissions:
   - Microsoft Graph: `Sites.Read.All` (Application permission)
   - Microsoft Graph: `Files.Read.All` (Application permission)
3. Grant admin consent for the permissions
4. Create a client secret

### Environment Variables
Set the following environment variables:

```bash
SHAREPOINT_URL=https://yourtenant.sharepoint.com
TENANT_ID=your-azure-tenant-id
CLIENT_ID=your-azure-app-client-id
CLIENT_SECRET=your-azure-app-client-secret
```

## Development

Install dependencies:
```bash
npm install
```

Build the server:
```bash
npm run build
```

For development with auto-rebuild:
```bash
npm run watch
```

## Testing

Test the server using the MCP Inspector:
```bash
npm run inspector
```

The Inspector provides a web interface to test all available tools and resources.

## Installation

### Claude Desktop Configuration

Add the server to your Claude Desktop configuration:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "sharepoint-mcp-server": {
      "command": "node",
      "args": ["/path/to/sharepoint-mcp-server/build/index.js"],
      "env": {
        "SHAREPOINT_URL": "https://yourtenant.sharepoint.com",
        "TENANT_ID": "your-azure-tenant-id",
        "CLIENT_ID": "your-azure-app-client-id",
        "CLIENT_SECRET": "your-azure-app-client-secret"
      }
    }
  }
}
```

### Global Installation

You can also install the server globally:

```bash
npm install -g .
```

Then use it directly:
```bash
sharepoint-mcp-server
```

## Architecture

The server implements a service-oriented architecture with clear separation of concerns:

- **Authentication Layer**: Handles OAuth2 token acquisition and refresh
- **Graph API Client**: Manages HTTP requests to Microsoft Graph API
- **Tool Handlers**: Process MCP tool requests and format responses
- **Resource Handlers**: Manage SharePoint site resources and metadata
- **Error Handling**: Comprehensive error management with proper MCP error codes

## Security Considerations

- Uses OAuth2 client credentials flow for secure authentication
- Tokens are automatically refreshed before expiration
- All API requests use HTTPS
- Client secrets should be stored securely and never committed to version control
- Application permissions require admin consent in Azure AD

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Verify Azure app registration and permissions
2. **Site Access**: Ensure the app has appropriate SharePoint permissions
3. **Network Issues**: Check firewall settings for Microsoft Graph API access

### Debug Mode

Set environment variable for detailed logging:
```bash
DEBUG=sharepoint-mcp-server
```

## Contributing

1. Follow TypeScript best practices
2. Maintain comprehensive error handling
3. Add tests for new functionality
4. Update documentation for API changes

## License

This project is licensed under the Mozilla Public License 2.0. See the [LICENSE](LICENSE) file for details.

## Contributing

We welcome contributions! Please follow these guidelines:

### Getting Started
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass (`npm test`)
6. Commit your changes (`git commit -m 'Add amazing feature'`)
7. Push to the branch (`git push origin feature/amazing-feature`)
8. Open a Pull Request

### Development Guidelines
- Follow TypeScript best practices and maintain type safety
- Implement comprehensive error handling with proper MCP error codes
- Add JSDoc comments for all public methods and classes
- Maintain the service-oriented architecture with clear separation of concerns
- Follow SOLID principles and keep functions focused and testable
- Update documentation for any API changes

### Code Style
- Use TypeScript strict mode
- Follow the existing code formatting and naming conventions
- Remove unused imports and variables
- Use descriptive variable and function names
- Prefer composition over inheritance

### Testing
- Add unit tests for new functionality
- Test error conditions and edge cases
- Ensure the basic test suite passes
- Test with real SharePoint environments when possible

### Documentation
- Update README.md for new features or configuration changes
- Add JSDoc comments for new public APIs
- Include examples for complex functionality
- Update the changelog for significant changes

## Changelog

### [0.1.0] - Initial Release
- Basic SharePoint integration via Microsoft Graph API
- Support for searching files across SharePoint
- Site listing and browsing capabilities
- Document library access and file content retrieval
- OAuth2 client credentials authentication
- MCP resource support for SharePoint sites
- Comprehensive error handling and logging

## Support

If you encounter issues or have questions:

1. Check the [troubleshooting section](#troubleshooting) in this README
2. Search existing [GitHub issues](https://github.com/sekops/sharepoint-mcp-server/issues)
3. Create a new issue with detailed information about your problem
4. Include relevant logs and configuration (without sensitive information)

## Acknowledgments

- Built with the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- Uses Microsoft Graph API for SharePoint integration
- Inspired by the MCP community and ecosystem
