# PowerPoint to Text Converter

A simple, browser-based web application that converts PowerPoint (.pptx) presentations to plain text files.

Visit [https://codehopperreddit.github.io/powerpointtotext/](https://codehopperreddit.github.io/fuzzymatcher/) to use the application immediately.

## Overview

This web application allows users to extract text content from PowerPoint presentations directly in their browser without requiring any server-side processing or installation. The app works entirely client-side, ensuring your presentations remain private and are never uploaded to any external server.

## Features

- **Browser-Based Conversion**: No server uploads required, all processing happens locally
- **Simple Drag & Drop Interface**: Easy-to-use interface for file selection
- **Text Preview**: View extracted text before downloading
- **Slide Organization**: Text is organized by slide number
- **Slide Notes Extraction**: Captures presenter notes if available
- **One-Click Download**: Save extracted text as a .txt file

## How It Works

PowerPoint (.pptx) files are actually ZIP archives containing XML files. This converter:

1. Unpacks the .pptx file using JSZip
2. Locates all slide XML files in the `ppt/slides/` directory
3. Extracts text content from each slide
4. Also extracts any presenter notes from the `ppt/notesSlides/` directory
5. Organizes the text by slide number
6. Allows you to preview and download the extracted text

## Usage Instructions

1. **Open the App**: Load the HTML file in any modern web browser
2. **Select a File**: Either drag and drop a .pptx file onto the upload area or click to browse
3. **Convert**: Click the "Convert to Text" button
4. **Preview**: Review the extracted text in the preview area
5. **Download**: Click "Download Text File" to save the extracted text

## Technical Requirements

- Modern web browser (Chrome, Firefox, Edge, Safari)
- JavaScript enabled
- Internet connection (only needed to load the JSZip library)

## Limitations

- Only supports .pptx format (newer PowerPoint format), not the older .ppt format
- Text extraction quality depends on how the PowerPoint was created
- May not perfectly preserve complex formatting or layout
- Does not extract text from images or charts within the presentation
- Text from SmartArt, tables, and other complex objects may have unpredictable formatting


## Privacy

This application processes all files locally in your browser. No data is sent to any server. Your presentations remain private.


## Credits

This application uses the following open-source libraries:
- [JSZip](https://stuk.github.io/jszip/) - For unpacking .pptx files

