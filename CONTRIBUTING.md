# Contributing to SharePoint MCP Server

Thank you for your interest in contributing to the SharePoint MCP Server! This document provides guidelines and information for contributors.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Making Changes](#making-changes)
- [Testing](#testing)
- [Submitting Changes](#submitting-changes)
- [Code Style Guidelines](#code-style-guidelines)
- [Architecture Guidelines](#architecture-guidelines)
- [Documentation](#documentation)

## Code of Conduct

This project follows a standard code of conduct. Please be respectful and professional in all interactions.

## Getting Started

### Prerequisites

- Node.js 18+ and npm
- TypeScript knowledge
- Basic understanding of Microsoft Graph API and SharePoint
- Familiarity with the Model Context Protocol (MCP)

### Development Setup

1. **Fork and Clone**
   ```bash
   git clone https://github.com/sekops/sharepoint-mcp-server.git
   cd sharepoint-mcp-server
   ```

2. **Install Dependencies**
   ```bash
   npm install
   ```

3. **Set Up Environment**
   ```bash
   cp .env.example .env
   # Edit .env with your Azure app credentials
   ```

4. **Build the Project**
   ```bash
   npm run build
   ```

5. **Run Tests**
   ```bash
   npm test
   ```

## Making Changes

### Branch Naming

Use descriptive branch names:
- `feature/add-new-tool` - for new features
- `fix/authentication-issue` - for bug fixes
- `docs/update-readme` - for documentation updates
- `refactor/improve-error-handling` - for code improvements

### Commit Messages

Follow conventional commit format:
```
type(scope): description

[optional body]

[optional footer]
```

Examples:
- `feat(tools): add file upload functionality`
- `fix(auth): handle token refresh edge case`
- `docs(readme): update installation instructions`

## Testing

### Running Tests

```bash
# Run basic tests
npm test

# Run with MCP Inspector for interactive testing
npm run inspector

# Watch mode for development
npm run watch
```

### Writing Tests

- Add tests for new functionality in the `test/` directory
- Test both success and error scenarios
- Include edge cases and boundary conditions
- Mock external dependencies when appropriate

### Test Structure

```javascript
// Example test structure
describe('Tool Name', () => {
  it('should handle valid input', async () => {
    // Test implementation
  });

  it('should throw error for invalid input', async () => {
    // Error case testing
  });
});
```

## Submitting Changes

### Pull Request Process

1. **Update Documentation**
   - Update README.md if adding new features
   - Add JSDoc comments for new public methods
   - Update CHANGELOG.md

2. **Create Pull Request**
   - Use a descriptive title
   - Include a detailed description of changes
   - Reference any related issues
   - Add screenshots for UI changes (if applicable)

3. **Pull Request Template**
   ```markdown
   ## Description
   Brief description of changes

   ## Type of Change
   - [ ] Bug fix
   - [ ] New feature
   - [ ] Breaking change
   - [ ] Documentation update

   ## Testing
   - [ ] Tests pass locally
   - [ ] Added tests for new functionality
   - [ ] Tested with MCP Inspector

   ## Checklist
   - [ ] Code follows style guidelines
   - [ ] Self-review completed
   - [ ] Documentation updated
   - [ ] No breaking changes (or clearly documented)
   ```

## Code Style Guidelines

### TypeScript Standards

- **Strict Mode**: Always use TypeScript strict mode
- **Type Safety**: Prefer explicit types over `any`
- **Interfaces**: Use interfaces for object shapes
- **Enums**: Use const assertions or string literal unions instead of enums when possible

### Naming Conventions

- **Classes**: PascalCase (`SharePointServer`)
- **Methods/Functions**: camelCase (`getAccessToken`)
- **Constants**: UPPER_SNAKE_CASE (`CLIENT_SECRET`)
- **Interfaces**: PascalCase with descriptive names (`GraphResponse`)

### Code Organization

```typescript
// 1. Imports (external libraries first, then internal)
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { customUtility } from "./utils.js";

// 2. Type definitions and interfaces
interface GraphResponse {
  value?: any[];
}

// 3. Constants
const DEFAULT_LIMIT = 10;

// 4. Class definition
class SharePointServer {
  // Private properties first
  private server: Server;
  
  // Constructor
  constructor() {
    // Implementation
  }
  
  // Public methods
  public async run(): Promise<void> {
    // Implementation
  }
  
  // Private methods
  private async getAccessToken(): Promise<string> {
    // Implementation
  }
}
```

### Error Handling

- Use proper MCP error codes
- Provide descriptive error messages
- Log errors appropriately
- Handle async operations with try-catch

```typescript
try {
  const result = await this.graphRequest(endpoint);
  return result;
} catch (error) {
  const errorMessage = error instanceof Error ? error.message : String(error);
  throw new McpError(ErrorCode.InternalError, `Operation failed: ${errorMessage}`);
}
```

## Architecture Guidelines

### Service-Oriented Architecture

The project follows a layered architecture:

1. **Transport Layer**: MCP protocol handling
2. **Service Layer**: Business logic and SharePoint operations
3. **Data Layer**: Microsoft Graph API interactions
4. **Authentication Layer**: OAuth2 token management

### SOLID Principles

- **Single Responsibility**: Each class/method has one clear purpose
- **Open/Closed**: Extend functionality without modifying existing code
- **Liskov Substitution**: Subtypes must be substitutable for base types
- **Interface Segregation**: Use specific interfaces rather than large ones
- **Dependency Inversion**: Depend on abstractions, not concretions

### Adding New Tools

When adding new MCP tools:

1. **Define the tool schema** in `setupToolHandlers()`
2. **Add the handler method** following naming convention `handle{ToolName}`
3. **Implement proper error handling** with MCP error codes
4. **Add comprehensive JSDoc comments**
5. **Update documentation** and tests

Example:
```typescript
{
  name: "new_tool",
  description: "Description of what the tool does",
  inputSchema: {
    type: "object",
    properties: {
      requiredParam: {
        type: "string",
        description: "Description of parameter",
      },
    },
    required: ["requiredParam"],
  },
}
```

### Adding New Resources

For new MCP resources:

1. **Define resource URI scheme** (e.g., `sharepoint://type/{id}`)
2. **Add to resource listing** in `ListResourcesRequestSchema` handler
3. **Implement resource reading** in `ReadResourceRequestSchema` handler
4. **Ensure proper MIME types** and metadata

## Documentation

### JSDoc Comments

Use comprehensive JSDoc for all public methods:

```typescript
/**
 * Retrieves detailed information about a SharePoint site
 * 
 * @param siteUrl - The SharePoint site URL
 * @returns Promise resolving to site information
 * @throws {McpError} When site URL is invalid or site is not accessible
 * 
 * @example
 * ```typescript
 * const siteInfo = await getSiteInfo('https://tenant.sharepoint.com/sites/mysite');
 * ```
 */
private async getSiteInfo(siteUrl: string): Promise<any> {
  // Implementation
}
```

### README Updates

When adding features:
- Update the features list
- Add new tool documentation
- Include usage examples
- Update installation instructions if needed

### Changelog

Follow [Keep a Changelog](https://keepachangelog.com/) format:

```markdown
## [Unreleased]

### Added
- New file upload tool for SharePoint documents

### Changed
- Improved error handling for authentication failures

### Fixed
- Fixed issue with special characters in file paths
```

## Getting Help

- **Issues**: Check existing GitHub issues or create a new one
- **Discussions**: Use GitHub Discussions for questions and ideas
- **Documentation**: Refer to the README and inline code comments

## License

By contributing, you agree that your contributions will be licensed under the Mozilla Public License 2.0.
