# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Open source preparation with Mozilla Public License 2.0
- Comprehensive contribution guidelines
- Enhanced documentation for open source community

## [0.1.0] - 2025-01-07

### Added
- Initial release of SharePoint MCP Server
- Microsoft Graph API integration for SharePoint access
- OAuth2 client credentials authentication with automatic token refresh
- MCP tools for SharePoint operations:
  - `search_files` - Search for files and documents across SharePoint
  - `list_sites` - List accessible SharePoint sites
  - `get_site_info` - Get detailed information about specific sites
  - `list_site_drives` - List document libraries in a site
  - `list_drive_items` - List files and folders in document libraries
  - `get_file_content` - Retrieve content of text files
- MCP resources for SharePoint sites with structured metadata
- Service-oriented architecture with clear separation of concerns
- Comprehensive error handling with proper MCP error codes
- TypeScript implementation with strict type safety
- Basic test suite for server functionality
- MCP Inspector integration for interactive testing
- Environment-based configuration
- Cross-platform support (Windows, macOS, Linux)

### Security
- Secure OAuth2 client credentials flow
- Automatic token refresh before expiration
- HTTPS-only API communication
- Environment variable-based secret management

### Documentation
- Comprehensive README with setup and usage instructions
- Azure app registration guide
- Installation instructions for Claude Desktop
- Architecture overview and security considerations
- Troubleshooting guide
- API documentation for all tools and resources

## Security Considerations

This project handles sensitive authentication credentials and accesses corporate SharePoint data. Please ensure:

- Client secrets are stored securely and never committed to version control
- Application permissions are granted only the minimum required access
- Regular security reviews of dependencies and authentication flows
- Proper network security when deploying in production environments

## Migration Guide

### From Pre-1.0 Versions

This is the initial public release. No migration is required.

## Support

For support and questions:
- Check the [README](README.md) troubleshooting section
- Review [GitHub Issues](https://github.com/sekops/sharepoint-mcp-server/issues)
- Create a new issue with detailed information about your problem
