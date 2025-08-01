# win-ocr

A Python package for performing Optical Character Recognition (OCR) on Windows using the built-in Windows.Media.Ocr API via PowerShell. This package provides high-accuracy, offline OCR similar to the performance seen in the Windows Photos app or native iPhone OCR—fast, accurate, and reliable for various fonts and layouts

## Motivation
The package enables programmatic access to Windows' native OCR capabilities without third-party libraries like Tesseract, which often underperform in accuracy compared to OS-integrated solutions. It avoids complex Python bindings (e.g., PyWinRT issues with modules not found or buffer handling) by leveraging PowerShell to directly interface with WinRT APIs. Suitable for command-line use or as a library, it supports file paths and optional Pillow Image objects, with compatibility for both Windows PowerShell 5.1 and PowerShell 7.

## Attribution
This package is heavily inspired by and adapted from the PsOcr PowerShell module by Tobias Weltner[](https://powershell.one/). The original code for loading WinRT types, handling async operations, and performing OCR is attributed to his work, licensed under MIT. We greatly appreciate his contributions, as the performance and simplicity were outstanding and enabled this Python wrapper. No legal issues arise from the adaptation, but full credit is given for the core PowerShell logic.

## How It Works
The package uses a Python class (`WinOCRSession`) to manage a persistent PowerShell session, initializing WinRT assemblies and the OCR engine once for efficiency. It supports:
- Processing image files (PNG, JPG, etc.) via paths.
- Optional Pillow Image objects (requires `pip install pillow`).
- Command-line interface for batch processing.
- Automatic detection of PowerShell version (5.1 or 7), with fallback to explicit selection.
- Offline operation on Windows 10/11 with the OCR language pack installed.

For .NET users, here's an untested C# example for a console app using the same API directly (add `Microsoft.Windows.SDK.Contracts` NuGet package in Visual Studio):

```csharp
using System;
using System.IO;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;

class Program
{
    static async Task Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: OCRApp.exe <image_path>");
            return;
        }

        string imagePath = args[0];
        string text = await ExtractTextFromImageAsync(imagePath);
        Console.WriteLine($"Extracted Text from {imagePath}:\n{text}");
    }

    static async Task<string> ExtractTextFromImageAsync(string imagePath)
    {
        try
        {
            // Load the image
            StorageFile file = await StorageFile.GetFileFromPathAsync(imagePath);

            // Open stream and decode bitmap
            using IRandomAccessStream fileStream = await file.OpenAsync(FileAccessMode.Read);
            BitmapDecoder decoder = await BitmapDecoder.CreateAsync(fileStream);
            SoftwareBitmap bitmap = await decoder.GetSoftwareBitmapAsync();

            // Initialize OCR engine
            OcrEngine engine = OcrEngine.TryCreateFromUserProfileLanguages();
            if (engine == null)
            {
                throw new Exception("OCR engine could not be created. Ensure language packs are installed.");
            }

            // Perform OCR
            OcrResult result = await engine.RecognizeAsync(bitmap);
            return result.Text;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            return null;
        }
    }
}
