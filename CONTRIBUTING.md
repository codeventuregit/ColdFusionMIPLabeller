# Contributing to ColdFusion MIP Labeller

Thank you for your interest in contributing to ColdFusion MIP Labeller! This document provides guidelines for contributing to the project.

## Getting Started

1. Fork the repository on GitHub
2. Clone your fork locally
3. Create a new branch for your feature or bug fix
4. Make your changes
5. Test your changes thoroughly
6. Submit a pull request

## Development Setup

### Prerequisites
- Visual Studio 2019 or later (or Visual Studio Code with C# extension)
- .NET Framework 4.8 SDK
- ColdFusion 2021+ (for testing integration)
- Microsoft Information Protection SDK

### Building the Project
```bash
dotnet restore ColdFusionMIPLabeller/ColdFusionMIPLabeller.csproj
dotnet build ColdFusionMIPLabeller/ColdFusionMIPLabeller.csproj --configuration Release
```

## Code Style

- Follow standard C# coding conventions
- Use meaningful variable and method names
- Add XML documentation comments for public methods
- Keep methods focused and single-purpose
- Handle exceptions appropriately with detailed error messages

## Testing

- Test all changes with actual ColdFusion integration
- Verify compatibility with different Office file formats
- Test error handling scenarios
- Include debug methods for troubleshooting

## Pull Request Process

1. Update documentation if you're changing functionality
2. Add or update XML comments for new public methods
3. Ensure your code builds without warnings
4. Test with ColdFusion integration
5. Update the README.md if needed
6. Submit your pull request with a clear description

## Reporting Issues

When reporting issues, please include:
- ColdFusion version
- .NET Framework version
- Error messages (full stack trace if available)
- Steps to reproduce the issue
- Sample code that demonstrates the problem

## Feature Requests

Feature requests are welcome! Please:
- Check existing issues first
- Provide a clear use case
- Explain how it would benefit ColdFusion developers
- Consider backward compatibility

## Code of Conduct

- Be respectful and inclusive
- Focus on constructive feedback
- Help others learn and grow
- Maintain a professional tone in all interactions

## License

By contributing, you agree that your contributions will be licensed under the MIT License.