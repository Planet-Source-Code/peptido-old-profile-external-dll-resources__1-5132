<div align="center">

## External DLL Resources


</div>

### Description

This code lets you read resources from a DLL external to an executable file

For example, let's say you have 20 BMPs, and 10 WAV file in your project,

and you don't want users to see them directly. You could put them in a

resource file, but you EXE file will be huge.

So, you can create a DLL with this resources, and then use this module to

read them
 
### More Info
 
Path to the DLL File, and resource name


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peptido \(old profile\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peptido-old-profile.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peptido-old-profile-external-dll-resources__1-5132/archive/master.zip)





### Source Code

```
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''    By: Peptido
''   Date: Dec 21 1999
''
''  Purpose: Reading resources from a DLL
''
''  Functions:
''
''   DrawDLLBitmap: Load a Bitmap Resource from the DLL and displays it
''    Parameters:
''      DLLPath: Path to the DLL file containing the resources
''      PicDesc: Name of the Bitmap Resource inside the DLL
''      hDC: Specifies where to Draw the bitmap
''      dstX: Optional. X coordinate specifying where to start drawing
''      dstY: Optional. Y coordinate specifying where to start drawing
''
''   DrawDLLIcon: Load an Icon Resource from the DLL and displays it
''    Parameters: Exactly the same as DrawDLLBitmap
''
''   LoadDLLString: Returns a String Resource in the DLL
''    Parameters:
''     DLLPath: Path to the DLL file containing the resources
''     StrNum: Number asigned to the String Resource
''
''   PlayDLLSound: Loads a Wave Resource from the DLL and plays it
''     DLLPath: Path to the DLL file containing the resources
''     WavDesc: Name of the Wave Resource inside the DLL
''
''
''  Known Bugs: None
''
''
''  Please send any comments, suggestions or bug reports to:
''    peptido@insideo.com.ar
''
'Structures Declaration
Private Type BITMAP
 bmType As Long
 bmWidth As Long
 bmHeight As Long
 bmWidthBytes As Long
 bmPlanes As Integer
 bmBitsPixel As Integer
 bmBits As Long
End Type
'Constant Declaration
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0
Private Const SRCCOPY = &HCC0020
'API Function Declaration
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Sub DrawDLLIcon(DLLPath As String, IconDesc As String, hDC As Long, Optional dstX As Long = 0, Optional dstY As Long = 0)
Dim hLibInst As Long
Dim hIcon As Long
hLibInst = LoadLibrary(DLLPath)
hIcon = LoadIcon(hLibInst, IconDesc)
Call DrawIcon(hDC, dstX, dstY, hIcon)
Call FreeLibrary(hLibInst)
End Sub
Public Sub DrawDLLBitmap(DLLPath As String, picDesc As String, hDC As Long, Optional dstX As Long = 0, Optional dstY As Long = 0)
Dim hLibInst As Long
Dim hdcMemory As Long
Dim hLoadedbitmap As Long
Dim hOldBitmap As Long
Dim bmpInfo As BITMAP
hLibInst = LoadLibrary(DLLPath)
hLoadedbitmap = LoadBitmap(hLibInst, picDesc)
Call GetObject(hLoadedbitmap, Len(bmpInfo), bmpInfo)
hdcMemory = CreateCompatibleDC(hDC)
hOldBitmap = SelectObject(hdcMemory, hLoadedbitmap)
Call BitBlt(hDC, dstX, dstY, bmpInfo.bmWidth, bmpInfo.bmHeight, hdcMemory, 0, 0, SRCCOPY)
Call SelectObject(hdcMemory, hOldBitmap)
Call DeleteObject(hLoadedbitmap)
Call DeleteDC(hdcMemory)
Call FreeLibrary(hLibInst)
End Sub
Public Sub PlayDLLSound(DLLPath As String, WavDesc As String)
Dim hLibInst As Long
hLibInst = LoadLibrary(DLLPath)
Call PlaySound(WavDesc, hLibInst, SND_RESOURCE Or SND_SYNC)
FreeLibrary (hLibInst)
End Sub
Public Function LoadDLLString(DLLPath As String, StrNum As Long) As String
Dim hLibInst As Long
Dim strTemp As String * 32768
Dim posTemp As Integer
hLibInst = LoadLibrary(DLLPath)
Call LoadString(hLibInst, StrNum, strTemp, Len(strTemp))
posTemp = InStr(strTemp, Chr$(0))
LoadDLLString = Left$(strTemp, posTemp - 1)
FreeLibrary (hLibInst)
End Function
```

