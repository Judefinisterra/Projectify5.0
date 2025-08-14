# Claude Configuration

This is the configuration file for Claude Code to help understand the Projectify 5.0 project.

## Project Overview
Excel add-in for financial modeling and business planning with AI assistance.

## Key Directories
- `src/taskpane/` - Main application logic
- `src/prompts/` - AI prompts and system instructions
- `assets/` - Icons and resources
- `docs/` - Documentation
- `dist/` - Build output

## Tech Stack
- JavaScript (ES6+)
- Office.js (Excel Add-in API)
- Node.js backend integration
- VBA macros for Excel automation

## Development Notes
- Main entry point: `src/taskpane/taskpane.js`
- Configuration: `src/taskpane/config.js`
- Build process uses webpack
- Testing integration exists in `src/taskpane/TestIntegration.js`

## Common Tasks
- Build: Check package.json for build scripts
- Test: npm test (if configured)
- Development server: Check package.json for dev scripts