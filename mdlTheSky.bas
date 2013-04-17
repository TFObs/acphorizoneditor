Attribute VB_Name = "TheSkyHoriz"
Option Explicit

Public Sub GetTheSkyHorizon()
Dim TShorizvals, AHEHorizVals(179)
Dim result
Dim x As Byte

result = MsgBox("Please click " & Chr(34) & "Copy" & Chr(34) & " in TheSky-Horizon Dialog." & vbCrLf & vbCrLf _
& "Then click OK to Continue or Cancel", vbOKCancel, "Open TheSky Horizon")

    If result = 1 Then
        'Get the Clipboard-Text and save it to the Horizon-Array
        TShorizvals = Split(Clipboard.GetText, Chr(10))

        For x = 0 To 179
            AHEHorizVals(x) = Trim(TShorizvals(x * 2))
        Next x
    
    Else: Exit Sub
    
    End If
    
    'Now Draw the Horizon
    For x = 0 To UBound(AHEHorizVals)
        frmMain.GridHoriz.TextMatrix(x + 1, 1) = AHEHorizVals(x)
    Next x
    
    frmMain.DrawHorizon

End Sub

Public Sub SaveTheSkyHorizon()
Dim TShorizvals$, azimval As Double
Dim result
Dim x As Byte

        TShorizvals = ""
        For x = 1 To 180
            If IsNumeric(frmMain.GridHoriz.TextMatrix(x, 1)) Then
                azimval = FormatNumber(frmMain.GridHoriz.TextMatrix(x, 1), 2)
                Else: azimval = FormatNumber("0.0", 2)
            End If
                
                TShorizvals = TShorizvals & Space(12 - Len(azimval)) & azimval & Chr(10)
                TShorizvals = TShorizvals & Space(12 - Len(azimval)) & azimval & Chr(10)
        Next x

        
    'Now Copy the Values to the CLipboard
    Clipboard.Clear
    Clipboard.SetText TShorizvals
    
    MsgBox "Please click " & Chr(34) & "Paste" & Chr(34) & " in TheSky-Horizon Dialog." & vbCrLf & vbCrLf _
& "To Copy the current Horizon to TheSky", vbOKOnly, "Save TheSky Horizon"

End Sub


Public Sub getTheSkyFile(ByVal Path As String)
Dim TSVals As String
Dim TShorizvals
Dim fs As New FileSystemObject
Dim infile As TextStream
Dim DescLength As Byte
Dim x As Integer

Set infile = fs.OpenTextFile(Path)

If check_TheSkyFile(Path) = 0 Then
    infile.Read (4)
    DescLength = CByte(Asc(infile.Read(1)))
    infile.Read (DescLength + 2)

    For x = 1 To 180
    frmMain.GridHoriz.TextMatrix(x, 1) = Format(Asc(infile.Read(1)) / 2, "0.0")
    infile.Read (1)
    Next x
Else
    infile.Read (22)
    frmMain.GridHoriz.TextMatrix(1, 1) = Format(Trim(infile.Read(8)), "0.0")
    For x = 2 To 180
        frmMain.GridHoriz.TextMatrix(x, 1) = Format(Trim(infile.Read(8)), "0.0")
        infile.Read (10)
    Next x
    
End If

Set infile = Nothing
Set fs = Nothing

End Sub


Public Sub saveTheSkyFile(ByVal Path As String, ByVal PathX As String)
Dim fs As New FileSystemObject
Dim outfile As TextStream
Dim TSXFile As TextStream
Dim desctext As String
Dim x
On Error Resume Next
Set outfile = fs.CreateTextFile(Path)
Set TSXFile = fs.CreateTextFile(PathX)

'======TheSky 6 Format===============
outfile.Write Chr(1) & Chr(0) & Chr(0) & Chr(0)
FloatWindow frmMain.hwnd, False
desctext = InputBox("Description for the Horizon: ", "Input a description")
If Len(desctext) > 255 Then desctext = Left(desctext, 255)
outfile.Write (Chr(Len(desctext)))
outfile.Write desctext & Chr(104) & Chr(1)

For x = 1 To 180
    outfile.Write Chr(Round(CDbl(frmMain.GridHoriz.TextMatrix(x, 1)) * 2, 0))
    outfile.Write Chr(Round(CDbl(frmMain.GridHoriz.TextMatrix(x, 1)) * 2, 0))
Next x

'======TheSkyX Format================
TSXFile.Write "   90.00|   90.00" & Chr(10)
TSXFile.Write "360" & Chr(10)

For x = 1 To 180
    TSXFile.Write Space(3) & Format(Round(CDbl(frmMain.GridHoriz.TextMatrix(x, 1)), 2), "00.00") & Chr(10)
    TSXFile.Write Space(3) & Format(Round(CDbl(frmMain.GridHoriz.TextMatrix(x, 1)), 2), "00.00") & Chr(10)
Next x

If Err.Number = 0 Then
    MsgBox "Files successfully created!", vbInformation, "Write successful!"
Else
    MsgBox Err.Number & " " & Err.Description & vbCrLf & "Error writing the Data, please write to file and try again!", vbCritical, "Error"
End If

TSXFile.Close: outfile.Close
Set outfile = Nothing: Set TSXFile = Nothing: Set fs = Nothing

If frmMain.mnuchkFloat.Checked Then FloatWindow frmMain.hwnd, True
End Sub

Function check_TheSkyFile(ByVal Path As String)
Dim fs As New FileSystemObject
Dim infile As TextStream
Set infile = fs.OpenTextFile(Path)
check_TheSkyFile = IIf(InStr(1, infile.Read(17), "|") = 9, 1, 0)
infile.Close
Set fs = Nothing
Set infile = Nothing

End Function
