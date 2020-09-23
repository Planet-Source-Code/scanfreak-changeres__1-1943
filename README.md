<div align="center">

## ChangeRes


</div>

### Description

This Function will change your Windows Resolution. It is very simple, and it does what most Resolution Change Functions don't do, it changes the the Bits Per Pixels as well as the Screen Width and Height.
 
### More Info
 
Dim RetValue As Integer

RetValue = ChangeRes(800, 600, 32)

1 = Resolution Successfully Changed

0 = Resolution Was Not Changed


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ScAnFrEaK](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scanfreak.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scanfreak-changeres__1-1943/archive/master.zip)

### API Declarations

```
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H60000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type
```


### Source Code

```
Function ChangeRes(Width As Single, Height As Single, BPP As Integer) As Integer
On Error GoTo ERROR_HANDLER
Dim DevM As DEVMODE, I As Integer, ReturnVal As Boolean, _
  RetValue, OldWidth As Single, OldHeight As Single, _
  OldBPP As Integer
  Call EnumDisplaySettings(0&, -1, DevM)
  OldWidth = DevM.dmPelsWidth
  OldHeight = DevM.dmPelsHeight
  OldBPP = DevM.dmBitsPerPel
  I = 0
  Do
    ReturnVal = EnumDisplaySettings(0&, I, DevM)
    I = I + 1
  Loop Until (ReturnVal = False)
  DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
  DevM.dmPelsWidth = Width
  DevM.dmPelsHeight = Height
  DevM.dmBitsPerPel = BPP
  Call ChangeDisplaySettings(DevM, 1)
  RetValue = MsgBox("Do You Wish To Keep Your Screen Resolution To " & Width & "x" & Height & " - " & BPP & " BPP?", vbQuestion + vbOKCancel, "Change Resolution Confirm:")
  If RetValue = vbCancel Then
    DevM.dmPelsWidth = OldWidth
    DevM.dmPelsHeight = OldHeight
    DevM.dmBitsPerPel = OldBPP
    Call ChangeDisplaySettings(DevM, 1)
    MsgBox "Old Resolution(" & OldWidth & " x " & OldHeight & ", " & OldBPP & " Bit) Successfully Restored!", vbInformation + vbOKOnly, "Resolution Confirm:"
    ChangeRes = 0
  Else
    ChangeRes = 1
  End If
  Exit Function
ERROR_HANDLER:
  ChangeRes = 0
End Function
```

