# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an Office Add-in for Excel, built using the Office.js framework. The add-in is named "Ecospold-Add-In" and provides functionality for exporting ecospold data. It's a task pane add-in that runs inside Excel and can interact with workbook data.

## Development Commands

### Building
- `npm run build` - Production build using webpack
- `npm run build:dev` - Development build
- `npm run watch` - Development build with file watching

### Development Server
- `npm run dev-server` - Start webpack dev server with HTTPS on port 3000
  - Uses self-signed certificates via `office-addin-dev-certs`
  - Dev server URL: https://localhost:3000/

### Debugging
- `npm start` - Start debugging the add-in in Excel (desktop by default)
  - Uses `manifest.xml` to sideload the add-in
  - App to debug: Excel (configurable in package.json `config.app_to_debug`)
  - App type: desktop (configurable in package.json `config.app_type_to_debug`)
- `npm stop` - Stop debugging session

### Linting and Formatting
- `npm run lint` - Check code with office-addin-lint
- `npm run lint:fix` - Auto-fix linting issues
- `npm run prettier` - Format code using prettier

### Manifest Operations
- `npm run validate` - Validate the manifest.xml file

### Authentication (M365 Account)
- `npm run signin` - Sign in to M365 account for debugging
- `npm run signout` - Sign out from M365 account

## Architecture

### Entry Points

The add-in has two main entry points defined in webpack.config.js:

1. **Taskpane** (`src/taskpane/taskpane.js` + `src/taskpane/taskpane.html`)
   - Main UI component shown in the Excel task pane
   - Contains the primary user interface and interaction logic
   - Exports a `run()` function for Excel operations
   - Contains a `download()` function for file download functionality

2. **Commands** (`src/commands/commands.js` + `src/commands/commands.html`)
   - Handles add-in commands that can be triggered from Excel ribbon buttons
   - Registers functions using `Office.actions.associate()`
   - Minimal UI, primarily for command execution

### Office.js Integration

- All code must wait for `Office.onReady()` before accessing Office APIs
- Excel operations use `Excel.run()` for batched execution with `context.sync()`
- The add-in targets Excel Workbook host (`<Host Name="Workbook"/>` in manifest)

### Build System

**Webpack Configuration** (webpack.config.js):
- Uses Babel for transpiling JavaScript
- Polyfills via core-js and regenerator-runtime for IE11 compatibility
- HTML files are processed through html-loader
- Assets (images) are copied to dist/assets/
- Manifest files are copied and URLs are replaced for production builds
  - Dev URL: https://localhost:3000/
  - Prod URL: https://www.contoso.com/ (update before production deployment)
- Generates separate bundles for taskpane and commands
- Development server uses HTTPS with auto-generated certificates

**Output Structure**:
- Built files go to `dist/` directory
- `dist/taskpane.html` - Main task pane UI
- `dist/commands.html` - Commands page
- `dist/manifest.xml` - Processed manifest with correct URLs
- `dist/assets/` - Icons and images

### Manifest (manifest.xml)

The manifest defines:
- Add-in ID, name, description, and provider
- Icon URLs (currently pointing to Azure storage: debugstoragebmgd239292.z1.web.core.windows.net)
- Source locations for taskpane and commands
- Excel ribbon integration (button in Home tab)
- Required permissions (ReadWriteDocument)

**Important**: Icon URLs and SourceLocation URLs in manifest.xml must be updated for production deployment. In dev mode, webpack copies the manifest unchanged. In production mode, it replaces localhost URLs with the production URL defined in webpack.config.js.

### Styling

- Task pane uses custom CSS in `src/taskpane/taskpane.css`
- Follow Office Add-in design patterns for consistency with Office UI

## Technical Details

### Browser Compatibility
- Targets last 2 versions of modern browsers and IE11 (see browserslist in package.json)
- Uses Babel preset-env for transpilation
- Includes polyfills for older environments

### Linting
- Uses `eslint-plugin-office-addins` with recommended rules
- Configuration in `.eslintrc.json`

### Current Implementation Notes

The taskpane.js currently includes:
- A `run()` function that demonstrates selecting a range and changing its fill color
- A `download()` function for file download functionality (work in progress)
  - Includes diagnostic logging for user agent, service worker support, localStorage
  - Creates blob URLs for file downloads
  - Note: File download behavior may vary depending on the Office host environment

### Development Workflow

1. Make changes to source files in `src/`
2. Use `npm run dev-server` for live development with hot reloading
3. Use `npm start` to sideload the add-in in Excel for testing
4. Before committing, run `npm run lint:fix` and `npm run validate`
5. For production deployment:
   - Update `urlProd` in webpack.config.js
   - Update icon URLs and other URLs in manifest.xml to production endpoints
   - Run `npm run build`
   - Deploy dist/ contents to your web server
