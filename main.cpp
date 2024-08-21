#include <windows.h>
#include <roapi.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Imaging.h>
#include <winrt/Windows.Storage.Streams.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Media.Ocr.h>
#include <winrt/Windows.Globalization.h>
#include <iostream>
#include <vector>
#include <fstream>
#include <string>
#include <locale>
#include <codecvt>

#include <gdiplus.h>
#pragma comment(lib, "gdiplus.lib")
Gdiplus::GdiplusStartupInput gdiplusStartupInput;
ULONG_PTR gdiplusToken;


int GetEncoderClsid(const WCHAR* format, CLSID* pClsid) {
    UINT num = 0;
    UINT size = 0;

    Gdiplus::ImageCodecInfo* pImageCodecInfo = NULL;
    Gdiplus::GetImageEncodersSize(&num, &size);
    if (size == 0) return -1;

    pImageCodecInfo = (Gdiplus::ImageCodecInfo*)(malloc(size));
    if (pImageCodecInfo == NULL) return -1;

    Gdiplus::GetImageEncoders(num, size, pImageCodecInfo);
    for (UINT j = 0; j < num; ++j) {
        if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0) {
            *pClsid = pImageCodecInfo[j].Clsid;
            free(pImageCodecInfo);
            return j;
        }
    }
    free(pImageCodecInfo);
    return -1;
}

std::wstring GetCurrentDirectory() {
    wchar_t buffer[MAX_PATH];
    if (GetCurrentDirectory(MAX_PATH, buffer)) {
        return std::wstring(buffer);
    }
    else {
        return L"";
    }
}

std::wstring CombinePath(const std::wstring& dir, const std::wstring& file) {
    std::wstring combined = dir;
    if (combined.back() != L'\\') {
        combined += L'\\';
    }
    combined += file;
    return combined;
}

void EnhanceImage(HBITMAP hBitmap, const std::wstring& outputPath) {
    // No need to initialize GDI+ here, as it's already done in main

    Gdiplus::Bitmap bitmap(hBitmap, NULL);
    int width = bitmap.GetWidth();
    int height = bitmap.GetHeight();

    Gdiplus::Bitmap enhancedBitmap(width * 2, height * 2, PixelFormat32bppARGB);
    Gdiplus::Graphics graphics(&enhancedBitmap);

    Gdiplus::ColorMatrix colorMatrix = {
        0.299f, 0.299f, 0.299f, 0.0f, 0.0f,
        0.587f, 0.587f, 0.587f, 0.0f, 0.0f,
        0.114f, 0.114f, 0.114f, 0.0f, 0.0f,
        0.0f, 0.0f, 0.0f, 1.0f, 0.0f,
        0.0f, 0.0f, 0.0f, 0.0f, 1.0f
    };

    Gdiplus::ImageAttributes imageAttributes;
    imageAttributes.SetColorMatrix(&colorMatrix, Gdiplus::ColorMatrixFlagsDefault, Gdiplus::ColorAdjustTypeBitmap);

    graphics.DrawImage(&bitmap, Gdiplus::Rect(0, 0, width, height), 0, 0, width, height, Gdiplus::UnitPixel, &imageAttributes);

    CLSID clsid;
    if (GetEncoderClsid(L"image/bmp", &clsid) >= 0) {
        enhancedBitmap.Save(outputPath.c_str(), &clsid, NULL);
        std::cout << "Enhanced image saved!" << std::endl;
    }
    else {
        std::cerr << "Failed to get encoder CLSID." << std::endl;
    }

    // Do not call GdiplusShutdown here
}

std::string WStringToString(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();

    int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.data(), static_cast<int>(wstr.size()), nullptr, 0, nullptr, nullptr);
    std::string str(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.data(), static_cast<int>(wstr.size()), &str[0], size_needed, nullptr, nullptr);
    return str;
}

HBITMAP CaptureScreen(int x, int y, int width, int height) {
    HDC hScreenDC = GetDC(NULL);
    HDC hMemoryDC = CreateCompatibleDC(hScreenDC);
    HBITMAP hBitmap = CreateCompatibleBitmap(hScreenDC, width, height);
    HBITMAP hOldBitmap = (HBITMAP)SelectObject(hMemoryDC, hBitmap);

    BitBlt(hMemoryDC, 0, 0, width, height, hScreenDC, x, y, SRCCOPY);
    hBitmap = (HBITMAP)SelectObject(hMemoryDC, hOldBitmap);

    DeleteDC(hMemoryDC);
    ReleaseDC(NULL, hScreenDC);

    return hBitmap;
}

bool SaveBitmapToFile(HBITMAP hBitmap, const std::wstring& filePath) {
    HDC hScreenDC = GetDC(NULL);
    if (!hScreenDC) {
        std::cerr << "Failed to get device context." << std::endl;
        return false;
    }

    BITMAPINFO bmpInfo = { 0 };
    bmpInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);

    if (GetDIBits(hScreenDC, hBitmap, 0, 0, nullptr, &bmpInfo, DIB_RGB_COLORS) == 0) {
        std::cerr << "Failed to get bitmap information." << std::endl;
        ReleaseDC(NULL, hScreenDC);
        return false;
    }

    bmpInfo.bmiHeader.biCompression = BI_RGB;
    bmpInfo.bmiHeader.biSizeImage = ((bmpInfo.bmiHeader.biWidth * bmpInfo.bmiHeader.biBitCount + 31) / 32) * 4 * bmpInfo.bmiHeader.biHeight;

    std::vector<BYTE> buffer(bmpInfo.bmiHeader.biSizeImage);

    if (GetDIBits(hScreenDC, hBitmap, 0, bmpInfo.bmiHeader.biHeight, buffer.data(), &bmpInfo, DIB_RGB_COLORS) == 0) {
        std::cerr << "Failed to get bitmap bits." << std::endl;
        ReleaseDC(NULL, hScreenDC);
        return false;
    }

    ReleaseDC(NULL, hScreenDC);

    BITMAPFILEHEADER fileHeader = { 0 };
    fileHeader.bfType = 0x4D42;
    fileHeader.bfOffBits = sizeof(BITMAPFILEHEADER) + bmpInfo.bmiHeader.biSize;
    fileHeader.bfSize = fileHeader.bfOffBits + bmpInfo.bmiHeader.biSizeImage;

    std::ofstream file(filePath, std::ios::binary);
    if (!file) {
        std::cerr << "Failed to open file for writing." << std::endl;
        return false;
    }

    file.write(reinterpret_cast<const char*>(&fileHeader), sizeof(fileHeader));
    file.write(reinterpret_cast<const char*>(&bmpInfo.bmiHeader), sizeof(bmpInfo.bmiHeader));
    file.write(reinterpret_cast<const char*>(buffer.data()), buffer.size());

    return true;
}

winrt::Windows::Graphics::Imaging::SoftwareBitmap LoadSoftwareBitmapFromFile(const std::wstring& filePath) {
    try {
        auto file = winrt::Windows::Storage::StorageFile::GetFileFromPathAsync(filePath).get();
        auto stream = file.OpenAsync(winrt::Windows::Storage::FileAccessMode::Read).get();
        auto decoder = winrt::Windows::Graphics::Imaging::BitmapDecoder::CreateAsync(stream).get();
        auto softwareBitmap = decoder.GetSoftwareBitmapAsync().get();
        return softwareBitmap;
    }
    catch (const winrt::hresult_error& ex) {
        std::cerr << "Failed to load SoftwareBitmap: " << WStringToString(ex.message().c_str()) << std::endl;
        return nullptr;
    }
}

// Function to delete a file
bool DeleteFile(const std::wstring& filePath) {
    if (DeleteFileW(filePath.c_str())) {
        std::wcout << L"Deleted file: " << filePath << std::endl;
        return true;
    }
    else {
        std::wcout << L"Failed to delete file: " << filePath << std::endl;
        return false;
    }
}

int main() {
    ::ShowWindow(::GetConsoleWindow(), SW_SHOW);
    winrt::init_apartment();

    // Initialize GDI+
    Gdiplus::GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

    while (true) {
        if ((GetAsyncKeyState('X') & 0x8000)) {
            int width = GetSystemMetrics(SM_CXSCREEN);
            int height = GetSystemMetrics(SM_CYSCREEN);
            HBITMAP hBitmap = CaptureScreen(0, 0, width, height);

            std::wstring currentDir = GetCurrentDirectory();
            std::wstring screenshotPath = CombinePath(currentDir, L"screenshot.bmp");
            std::wstring enhancedPath = CombinePath(currentDir, L"enhanced_screenshot.bmp");

            if (SaveBitmapToFile(hBitmap, screenshotPath)) {
                std::wcout << L"Original bitmap saved to file: " << screenshotPath << std::endl;

                EnhanceImage(hBitmap, enhancedPath);
                std::wcout << L"Enhanced bitmap saved to file: " << enhancedPath << std::endl;

                auto softwareBitmap = LoadSoftwareBitmapFromFile(enhancedPath);
                if (softwareBitmap) {
                    std::wcout << L"Loaded enhanced image successfully. Performing OCR..." << std::endl;

                    // Reinitialize the OCR engine in each iteration
                    auto language = winrt::Windows::Globalization::Language(L"en");
                    auto ocrEngine = winrt::Windows::Media::Ocr::OcrEngine::TryCreateFromLanguage(language);
                    if (!ocrEngine) {
                        std::cerr << "Failed to create OCR engine for language: en" << std::endl;
                        return -1;
                    }

                    // Perform OCR
                    auto ocrResultAsync = ocrEngine.RecognizeAsync(softwareBitmap);
                    auto ocrResult = ocrResultAsync.get(); // Get the result synchronously

                    // Print recognized text to console
                    auto lines = ocrResult.Lines();
                    for (auto&& line : lines) {
                        std::wcout << line.Text().c_str() << std::endl;
                    }
                    std::wcout << L"Text successfully recognized and printed to console." << std::endl;
                }
                else {
                    std::cerr << "Failed to load SoftwareBitmap from file." << std::endl;
                }

                // Delete the screenshot and enhanced image files
                DeleteFile(screenshotPath);
                DeleteFile(enhancedPath);
            }
            else {
                std::cerr << "Failed to save bitmap to file." << std::endl;
            }

            DeleteObject(hBitmap);

            std::wcout << L"Waiting for the next iteration..." << std::endl;
            Sleep(5000); // Wait for 5 seconds before the next iteration
        }
    }

    // Shutdown GDI+
    Gdiplus::GdiplusShutdown(gdiplusToken);

    std::wcout << L"Exiting the loop. Program terminated." << std::endl;
    return 0;
}