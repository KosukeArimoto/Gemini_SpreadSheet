# Gemini SpreadSheet

Google Apps Script project for Gemini integration with Google Sheets.

## Setup

### Prerequisites
- Node.js (installed via Homebrew)
- clasp (Google Apps Script CLI)

### Installation

1. Clone or download this repository
2. Install dependencies:
```bash
npm install
```

## Development

### Available Commands

- `npm run push` - Push local changes to Google Apps Script
- `npm run pull` - Pull changes from Google Apps Script to local
- `npm run open` - Open the project in Google Apps Script editor
- `npm run deploy` - Push and deploy the script
- `npm run watch` - Watch for changes and auto-push

### Workflow

1. Edit files locally in your preferred editor
2. Push changes to GAS:
```bash
npm run push
```

3. Or use watch mode for automatic syncing:
```bash
npm run watch
```

## Project Structure

- `*.js` - Google Apps Script files
- `appsscript.json` - GAS manifest file
- `.clasp.json` - clasp configuration
- `tsconfig.json` - TypeScript configuration (for editor support)

## Files

### Core Files
- `setGrobalVariables.js` - Global variables and constants setup
- `commonHelpers.js` - **NEW** Shared utility functions used across multiple files
- `callGenerativeAI.js` - Generative AI API integration (Gemini & OpenAI)
- `setUserCredencials.js` - User credentials management
- `userInterface.js` - UI menu definitions

### Feature Files
- `generateCategories.js` - Category generation and feedback logic
- `generateSlides.js` - Slide generation with batch processing
- `generateSlidesTR.js` - Slide generation (TR version)
- `generateRowImages_batch.js` - **NEW** Row-by-row image generation with batch processing
- `forTokairika.js` - Tokairika specific functions
- `memo.gs.js` - Memo functions

### Shared Functions in commonHelpers.js

The following utility functions have been consolidated into `commonHelpers.js`:
- `_extractFolderIdFromUrl()` - Extract Google Drive folder ID from URL
- `_parseNumberRangeString()` - Parse number ranges (e.g., "1-5, 10, 15-20")
- `_parseColumnRangeString()` - Parse column ranges (e.g., "A, C, E-G")
- `_columnToIndex()` - Convert column letter to index
- `parseMarkdownTable_()` - Parse markdown table to 2D array
- `_replacePrompts()` - Replace placeholders in prompts
- `extractGoogleDriveId_()` - Extract Google Drive ID from URL
- `_showSetupCompletionDialog()` - Show setup completion dialog
- `stopTriggers_()` - Stop triggers for a specific function
