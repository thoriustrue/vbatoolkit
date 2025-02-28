# Excel VBA Toolkit

A browser-based toolkit for working with Excel VBA files. This application allows you to remove VBA password protection, extract VBA code, and perform other operations on Excel files with macros.

## Features

- **VBA Password Removal**: Remove password protection from VBA projects in Excel files
- **VBA Code Extraction**: Extract all VBA code modules from Excel files
- **Sheet Protection Removal**: Remove worksheet and workbook protection
- **Trust Settings Management**: Configure files to auto-enable macros and external links
- **100% Client-Side Processing**: All operations happen in your browser - no data is sent to any server

## Supported File Formats

- Excel 2007-2022 files (.xlsm, .xls, .xlsb)
- Files with VBA macros and password protection

## Technical Overview

This application is built with:

- React for the UI
- TypeScript for type safety
- SheetJS for Excel file parsing
- JSZip for ZIP file manipulation
- TailwindCSS for styling

## Project Structure

```
src/
├── components/         # React components
│   ├── ErrorLogger.tsx # Error handling system
│   ├── FileUploader.tsx # File upload component
│   ├── LogViewer.tsx   # Log display component
│   └── ...
├── utils/              # Utility functions
│   ├── vbaCodeExtractor/ # VBA code extraction
│   │   ├── index.ts    # Main extraction logic
│   │   ├── types.ts    # Type definitions
│   │   └── ...
│   ├── vbaPasswordRemover.ts # Password removal logic
│   ├── fileUtils.ts    # File handling utilities
│   └── ...
├── types.ts            # Global type definitions
├── App.tsx             # Main application component
└── main.tsx            # Application entry point
```

## Development

### Prerequisites

- Node.js 16+
- npm or yarn

### Setup

1. Clone the repository
2. Install dependencies:
   ```
   npm install
   ```
3. Start the development server:
   ```
   npm run dev
   ```

### Building for Production

```
npm run build
```

The build output will be in the `dist` directory.

### Deployment

The application is configured for deployment to GitHub Pages:

```
npm run deploy
```

## Error Handling

The application includes a comprehensive error handling system:

- Client-side error boundary to catch and display React errors
- Detailed logging system for operation progress and errors
- Error recovery options for common failure scenarios

## Security Considerations

- All processing happens client-side for maximum privacy
- No data is sent to any server
- The application is designed for legitimate use cases only, such as:
  - Recovering your own VBA code when you've forgotten the password
  - Extracting VBA code for backup or migration
  - Removing protection from your own files

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Updating the Changelog

When making significant changes, please update the changelog:

```bash
# Add a new version with changes
npm run changelog -- --version "1.2.0" --add "New feature" --fix "Bug fix" --change "Changed behavior"
```

### Setting Up the Pre-commit Hook

To be reminded to update the changelog when committing changes:

```bash
# Make the script executable
chmod +x scripts/pre-commit.js

# Create a symlink in the git hooks directory
ln -s ../../scripts/pre-commit.js .git/hooks/pre-commit
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Disclaimer

This tool should only be used on Excel files that you own or have explicit permission to modify. Unauthorized access to protected files may violate applicable laws. 