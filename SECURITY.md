# Security Policy

## Supported Versions

We release patches for security vulnerabilities. Which versions are eligible for receiving such patches depends on the CVSS v3.0 Rating:

| Version | Supported          |
| ------- | ------------------ |
| 0.1.x   | :white_check_mark: |

## Reporting a Vulnerability

The SharePoint MCP Server team takes security bugs seriously. We appreciate your efforts to responsibly disclose your findings, and will make every effort to acknowledge your contributions.

### How to Report a Security Vulnerability

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please report them via email to: [security@sekops.ch] 

You should receive a response within 48 hours. If for some reason you do not, please follow up via email to ensure we received your original message.

Please include the following information in your report:

- Type of issue (e.g. buffer overflow, SQL injection, cross-site scripting, etc.)
- Full paths of source file(s) related to the manifestation of the issue
- The location of the affected source code (tag/branch/commit or direct URL)
- Any special configuration required to reproduce the issue
- Step-by-step instructions to reproduce the issue
- Proof-of-concept or exploit code (if possible)
- Impact of the issue, including how an attacker might exploit the issue

This information will help us triage your report more quickly.

## Security Considerations

### Authentication and Authorization

- **OAuth2 Client Credentials**: The server uses OAuth2 client credentials flow for authentication with Microsoft Graph API
- **Token Management**: Access tokens are automatically refreshed and stored in memory only
- **Least Privilege**: Configure Azure app permissions with minimal required access
- **Admin Consent**: Application permissions require admin consent in Azure AD

### Data Protection

- **In-Transit**: All communications with Microsoft Graph API use HTTPS
- **At-Rest**: No persistent storage of authentication tokens or SharePoint data
- **Logging**: Sensitive information is not logged (tokens, credentials, file contents)
- **Environment Variables**: Credentials are managed through environment variables

### Network Security

- **HTTPS Only**: All external API calls use HTTPS
- **No Inbound Connections**: Server operates as a client, no inbound network connections
- **Firewall**: Ensure outbound HTTPS access to Microsoft Graph API endpoints

### Configuration Security

- **Secret Management**: Store client secrets securely, never in source code
- **Environment Isolation**: Use separate Azure app registrations for different environments
- **Access Review**: Regularly review and audit Azure app permissions
- **Credential Rotation**: Implement regular rotation of client secrets

### Deployment Security

- **Container Security**: If using containers, follow container security best practices
- **File Permissions**: Ensure proper file permissions on configuration files
- **Process Isolation**: Run the server with minimal required privileges
- **Monitoring**: Implement logging and monitoring for security events

## Security Best Practices

### For Developers

1. **Input Validation**: Always validate and sanitize user inputs
2. **Error Handling**: Don't expose sensitive information in error messages
3. **Dependencies**: Keep dependencies updated and monitor for vulnerabilities
4. **Code Review**: All security-related changes require thorough review
5. **Testing**: Include security testing in the development process

### For Administrators

1. **Azure App Configuration**:
   - Use application permissions, not delegated permissions
   - Grant only the minimum required permissions
   - Regularly review and audit permissions
   - Enable conditional access policies

2. **Environment Setup**:
   - Use separate Azure apps for development, staging, and production
   - Implement proper secret management (Azure Key Vault, etc.)
   - Enable audit logging in Azure AD
   - Monitor for unusual authentication patterns

3. **Network Configuration**:
   - Restrict network access where possible
   - Use VPN or private networks for sensitive environments
   - Implement proper firewall rules
   - Monitor network traffic for anomalies

### For End Users

1. **Claude Desktop Configuration**:
   - Store configuration files securely
   - Don't share configuration files containing credentials
   - Use environment-specific configurations
   - Regularly review active MCP servers

## Known Security Limitations

1. **Client Credentials Flow**: Uses application permissions which have broad access
2. **Memory Storage**: Tokens are stored in memory and could be accessed by memory dumps
3. **Process Access**: Any process with sufficient privileges could access environment variables
4. **Network Monitoring**: Network traffic could be monitored if HTTPS is compromised

## Security Updates

Security updates will be released as patch versions and documented in the [CHANGELOG](CHANGELOG.md). 

Subscribe to releases on GitHub to be notified of security updates.

## Compliance

This project is designed to work with:

- **Microsoft 365 Commercial**: Standard commercial tenants
- **Microsoft 365 Government**: GCC, GCC High, and DoD environments
- **SharePoint Server**: On-premises deployments

Ensure your deployment meets your organization's compliance requirements.

## Third-Party Dependencies

We regularly monitor our dependencies for security vulnerabilities using:

- npm audit
- GitHub Dependabot
- Manual security reviews

## Contact

For security-related questions or concerns, please contact: [security@sekops.ch]

For general questions, please use the GitHub issues or discussions.
