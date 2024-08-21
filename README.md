# Strange Printer OCR Tool

## Overview

This tool captures the screen, enhances the captured image, performs Optical Character Recognition (OCR) on it, and then deletes the temporary files. It uses the Windows API, GDI+, and WinRT for various tasks.

## Prerequisites

- **Windows Development Kit**: Ensure you have the Windows SDK installed.
- **Visual Studio**: Recommended for building and running the application.
- **GDI+**: Used for image manipulation.
- **WinRT**: Used for OCR functionality.

## Building the Project

1. **Create a New Project**:
   - Open Visual Studio.
   - Create a new "Console Application" project.

2. **Add the Code**:
   - Replace the contents of the main source file with the provided code.

3. **Configure Project Settings**:
   - **Link Against GDI+**:
     - Go to Project Properties -> Linker -> Input -> Additional Dependencies.
     - Add `gdiplus.lib`.
   - **Include Directories**:
     - Ensure you have the Windows SDK directories added to the Include Directories.

4. **Compile the Project**:
   - Build the project using Visual Studio.

## Running the Tool

1. **Run the Executable**:
   - Execute the built application. 

2. **Usage**:
   - Press the `X` key to trigger the screen capture.
   - The tool will capture the screen, save and enhance the image, perform OCR, print the recognized text to the console, and delete the temporary files.

## Notes

- **Initialization and Cleanup**:
  - The tool initializes GDI+ and WinRT in the `main()` function and shuts down GDI+ when the program exits.

- **File Handling**:
  - The tool saves the captured and enhanced images to the current working directory and deletes them after processing.

## Example

To test the functionality:
1. Build the project in Visual Studio.
2. Run the executable.
3. Press `X` to capture and process the screen.
