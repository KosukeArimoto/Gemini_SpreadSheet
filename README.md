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

- `callGenerativeAI.js` - Generative AI integration
- `generateCategories.js` - Category generation logic
- `generateSlides.js` - Slide generation
- `generateSlidesTR.js` - Slide generation (TR version)
- `forTokairika.js` - Tokairika specific functions
- `memo.gs.js` - Memo functions
- `setGrobalVariables.js` - Global variables setup
- `setUserCredencials.js` - User credentials management
- `userInterface.js` - UI components
