' Author: iBug
' Converts all PPTX to PDF in current directory

Option Explicit

Dim Shell, FS
Set Shell = CreateObject("WScript.Shell")
Set FS = CreateObject("Scripting.FileSystemObject")
Dim PowerPoint

' Dim SupportedExtension : SupportedExtension = "pptx"
Dim SupportedExtensions : SupportedExtensions = Array("pptx", "ppt")
Dim SupportedExtension

Function Conv(FileName, SupportedExtension)
  Dim PPT, Range, SaveName
  Set PPT = PowerPoint.Presentations.Open(FileName)
  Set Range = PPT.PrintOptions.Ranges.Add(1, 1)
  SaveName = Replace(FileName, "." & SupportedExtension, ".pdf")
  PPT.ExportAsFixedFormat SaveName, 2, 2, 0, 2, 4, 0, Range, 1, False, False, False, False, False
  PPT.Close
End Function

Sub ConvAll(Dir)
  Dim Item
  For Each Item In Dir.Files
    For Each SupportedExtension In SupportedExtensions
      If LCase(FS.GetExtensionName(Item.Path)) = SupportedExtension Then
        Call Conv(Item.Path, SupportedExtension)
      End If
    Next
  Next
  For Each Item In Dir.SubFolders
    ConvAll Item
  Next
End Sub

Set PowerPoint = CreateObject("PowerPoint.Application")
PowerPoint.Visible = True
ConvAll FS.GetFolder(".")
PowerPoint.Quit