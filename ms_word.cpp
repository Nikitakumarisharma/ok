#include <iostream>
#include <fstream>
#include <string>
#include <windows.h>

using namespace std;

int main()
{
    string wordFile = "C:\\Users\\nikit\\OneDrive\\Documents\\Nikita_123.docx";
    string pptFile = "C:\\Users\\nikit\\OneDrive\\Documents\\Presentation1.pptx";

    // Convert file paths to wide-character strings
    wstring wordFileW(wordFile.begin(), wordFile.end());
    wstring pptFileW(pptFile.begin(), pptFile.end());

    // Open Word document
    ShellExecuteW(NULL, L"open", wordFileW.c_str(), NULL, NULL, SW_SHOWNORMAL);

    // Wait for Word to open
    Sleep(5000);

    // Copy text from Word
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event('A', 0, 0, 0);
    keybd_event('A', 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event('C', 0, 0, 0);
    keybd_event('C', 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

    // Open PowerPoint presentation
    ShellExecuteW(NULL, L"open", pptFileW.c_str(), NULL, NULL, SW_SHOWNORMAL);

    // Wait for PowerPoint to open
    Sleep(5000);

    // Paste text into PowerPoint
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event('V', 0, 0, 0);
    keybd_event('V', 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

    return 0;
}



