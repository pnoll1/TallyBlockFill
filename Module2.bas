Attribute VB_Name = "Module2"
Sub Filltallyblock()
'cell of lift must be clicked before starting macro
'Tally sheet workbooks names must match names used below

'Declare variables
Dim x As Integer
Dim y As Integer
Dim wb As String
x = ActiveCell.Row
y = ActiveCell.Column

'Variables for file system objects
Dim path As String
path = ThisWorkbook.path & "\"
Dim fso As Object
Set fso = VBA.CreateObject("scripting.FileSystemobject")

'determine if for erect or disman
'indented if block determine which lift this is by searching description column for keyword
If ActiveSheet.Name = "ERECT" Then
    e = "ERECT"
    If InStr(ActiveSheet.Cells(x, y + 1), "Tower") >= 1 Then
        wb = "Tally Sheet for Tower"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Slewing") >= 1 Then
        wb = "Tally Sheet for Slewing"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Machine Deck") >= 1 Then
        wb = "Tally Sheet for Machine Deck"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Counterjib") >= 1 Then
        wb = "Tally Sheet for Counterjib"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Hoist") >= 1 Then
        wb = "Tally Sheet for Hoist"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Inner Jib") >= 1 Then
        wb = "Tally Sheet for Inner Jib"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Outer Jib") >= 1 Then
        wb = "Tally Sheet for Outer Jib"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Cwt") >= 1 Then
        wb = "Tally Sheet for Counterweight"
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    End If
ElseIf ActiveSheet.Name = "DISMAN" Then
    e = "DISMAN"
    If InStr(ActiveSheet.Cells(x, y + 1), "Tower" Or "Slewing" Or "Machine Deck") > 1 Then
        wb = "Tally Sheet for Tower" & e
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Counterjib") >= 1 Then
        wb = "Tally Sheet for Counterjib" & e
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Hoist") >= 1 Then
        wb = "Tally Sheet for Hoist" & e
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Inner Jib") >= 1 Then
        wb = "Tally Sheet for Inner Jib" & e
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Outer Jib") >= 1 Then
        wb = "Tally Sheet for Outer Jib" & e
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    ElseIf InStr(ActiveSheet.Cells(x, y + 1), "Cwt") >= 1 Then
        wb = "Tally Sheet for Counterweight" & e
        fso.CopyFile "S:\Sicklesteel Cranes\Engineering\Misc\Tower Cranes\Tally Sheet for ACAD.xlsx", path
        Workbooks.Open (path & "Tally Sheet for ACAD.xlsx")
        ActiveWorkbook.SaveAs Filename:=path & wb & ".xlsx"
    End If
End If

'Figure out which tally sheet worksheet to write to
If ThisWorkbook.Worksheets(e).Range("Boom_Config").Value = "SH" Then
    ws = "Main Boom (Head)"
    Workbooks(wb).Worksheets(ws).Activate
    Range("Load").Formula = ThisWorkbook.Worksheets(e).Cells(x, y + 8).Value 'Load
    Range("Capacity").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 11).Value 'Capacity
    Range("Max_Rad").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 13).Value 'Radius
    Range("Name").Value = ThisWorkbook.Worksheets(e).Range("Name").Value 'Crane Name
    Range("Tonnage").Value = ThisWorkbook.Worksheets(e).Range("Tonnage").Value 'Tonnage
    Range("Main_Len").Value = ThisWorkbook.Worksheets(e).Range("Main_Len").Value 'Main Boom Length
    Range("CWT").Value = ThisWorkbook.Worksheets(e).Range("CWT").Value 'Counterweight
    Range("Block").Value = ThisWorkbook.Worksheets(e).Range("Block").Value 'Block
    Range("Ball").Value = ThisWorkbook.Worksheets(e).Range("Ball").Value 'Ball
    Range("Rigging").Value = ThisWorkbook.Worksheets(e).Range("Rigging").Value 'Rigging
    ActiveWorkbook.Save
    
ElseIf ThisWorkbook.Worksheets(e).Range("Boom_Config").Value = "SHSL" Then
    ws = "Main Boom (Head)"
    Workbooks(wb).Worksheets(ws).Activate
    Range("Load").Formula = ThisWorkbook.Worksheets(e).Cells(x, y + 8).Value 'Load
    Range("Capacity").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 11).Value 'Capacity
    Range("Max_Rad").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 13).Value 'Radius
    Range("Name").Value = ThisWorkbook.Worksheets(e).Range("Name").Value 'Crane Name
    Range("Tonnage").Value = ThisWorkbook.Worksheets(e).Range("Tonnage").Value 'Tonnage
    Range("Main_Len").Value = ThisWorkbook.Worksheets(e).Range("Main_Len").Value 'Main Boom Length
    Range("CWT").Value = ThisWorkbook.Worksheets(e).Range("CWT").Value 'Counterweight
    Range("Block").Value = ThisWorkbook.Worksheets(e).Range("Block").Value 'Block
    Range("Ball").Value = ThisWorkbook.Worksheets(e).Range("Ball").Value 'Ball
    Range("Rigging").Value = ThisWorkbook.Worksheets(e).Range("Rigging").Value 'Rigging
    Range("Super_Lift").Value = Yes 'superlift
    ActiveWorkbook.Save
ElseIf ThisWorkbook.Worksheets(e).Range("Boom_Config").Value = "SA" Then
    ws = "Swing Away"
    Workbooks(wb).Worksheets(ws).Activate
    Range("Load").Formula = ThisWorkbook.Worksheets(e).Cells(x, y + 8).Value 'Load
    Range("Capacity").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 11).Value 'Capacity
    Range("Max_Rad").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 13).Value 'Radius
    Range("Name").Value = ThisWorkbook.Worksheets(e).Range("Name").Value 'Crane Name
    Range("Tonnage").Value = ThisWorkbook.Worksheets(e).Range("Tonnage").Value 'Tonnage
    Range("Main_Len").Value = ThisWorkbook.Worksheets(e).Range("Main_Len").Value 'Main Boom Length
    Range("CWT").Value = ThisWorkbook.Worksheets(e).Range("CWT").Value 'Counterweight
    Range("Block").Value = ThisWorkbook.Worksheets(e).Range("Block").Value 'Block
    Range("Ball").Value = ThisWorkbook.Worksheets(e).Range("Ball").Value 'Ball
    Range("Rigging").Value = ThisWorkbook.Worksheets(e).Range("Rigging").Value 'Rigging
    Cells(31, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Len").Value 'Jib Length
    Cells(32, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Angle").Value 'Jib_Angle
    ActiveWorkbook.Save
ElseIf ThisWorkbook.Worksheets(e).Range("Boom_Config").Value = "SASL" Then
    ws = "Swing Away"
    Workbooks(wb).Worksheets(ws).Activate
    Range("Load").Formula = ThisWorkbook.Worksheets(e).Cells(x, y + 8).Value 'Load
    Range("Capacity").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 11).Value 'Capacity
    Range("Max_Rad").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 13).Value 'Radius
    Range("Name").Value = ThisWorkbook.Worksheets(e).Range("Name").Value 'Crane Name
    Range("Tonnage").Value = ThisWorkbook.Worksheets(e).Range("Tonnage").Value 'Tonnage
    Range("Main_Len").Value = ThisWorkbook.Worksheets(e).Range("Main_Len").Value 'Main Boom Length
    Range("CWT").Value = ThisWorkbook.Worksheets(e).Range("CWT").Value 'Counterweight
    Range("Block").Value = ThisWorkbook.Worksheets(e).Range("Block").Value 'Block
    Range("Ball").Value = ThisWorkbook.Worksheets(e).Range("Ball").Value 'Ball
    Range("Rigging").Value = ThisWorkbook.Worksheets(e).Range("Rigging").Value 'Rigging
    Cells(31, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Len").Value 'Jib Length
    Cells(32, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Angle").Value 'Jib_Angle
    Range("Super_Lift").Value = Yes 'superlift
    ActiveWorkbook.Save
ElseIf ThisWorkbook.Worksheets(e).Range("Boom_Config").Value = "SF" Then
    ws = "Fixed Jib"
    Workbooks(wb).Worksheets(ws).Activate
    Range("Load").Formula = ThisWorkbook.Worksheets(e).Cells(x, y + 8).Value 'Load
    Range("Capacity").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 11).Value 'Capacity
    Range("Max_Rad").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 13).Value 'Radius
    Range("Name").Value = ThisWorkbook.Worksheets(e).Range("Name").Value 'Crane Name
    Range("Tonnage").Value = ThisWorkbook.Worksheets(e).Range("Tonnage").Value 'Tonnage
    Range("Main_Len").Value = ThisWorkbook.Worksheets(e).Range("Main_Len").Value 'Main Boom Length
    Range("CWT").Value = ThisWorkbook.Worksheets(e).Range("CWT").Value 'Counterweight
    Range("Block").Value = ThisWorkbook.Worksheets(e).Range("Block").Value 'Block
    Range("Ball").Value = ThisWorkbook.Worksheets(e).Range("Ball").Value 'Ball
    Range("Rigging").Value = ThisWorkbook.Worksheets(e).Range("Rigging").Value 'Rigging
    Cells(32, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Len").Value 'Jib Length
    Cells(33, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Angle").Value 'Jib_Angle
    ActiveWorkbook.Save
ElseIf ThisWorkbook.Worksheets(e).Range("Boom_Config").Value = "SFSL" Then
    ws = "Fixed Jib"
    Workbooks(wb).Worksheets(ws).Activate
    Range("Load").Formula = ThisWorkbook.Worksheets(e).Cells(x, y + 8).Value 'Load
    Range("Capacity").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 11).Value 'Capacity
    Range("Max_Rad").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 13).Value 'Radius
    Range("Name").Value = ThisWorkbook.Worksheets(e).Range("Name").Value 'Crane Name
    Range("Tonnage").Value = ThisWorkbook.Worksheets(e).Range("Tonnage").Value 'Tonnage
    Range("Main_Len").Value = ThisWorkbook.Worksheets(e).Range("Main_Len").Value 'Main Boom Length
    Range("CWT").Value = ThisWorkbook.Worksheets(e).Range("CWT").Value 'Counterweight
    Range("Block").Value = ThisWorkbook.Worksheets(e).Range("Block").Value 'Block
    Range("Ball").Value = ThisWorkbook.Worksheets(e).Range("Ball").Value 'Ball
    Range("Rigging").Value = ThisWorkbook.Worksheets(e).Range("Rigging").Value 'Rigging
    Cells(32, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Len").Value 'Jib Length
    Cells(33, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Angle").Value 'Jib_Angle
    Range("Super_Lift").Value = Yes 'superlift
    ActiveWorkbook.Save
ElseIf ThisWorkbook.Worksheets(e).Range("Boom_Config").Value = "SW" Then
    ws = "Luffing Jib"
    Workbooks(wb).Worksheets(ws).Activate
    Range("Load").Formula = ThisWorkbook.Worksheets(e).Cells(x, y + 8).Value 'Load
    Range("Capacity").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 11).Value 'Capacity
    Range("Max_Rad").Value = ThisWorkbook.Worksheets(e).Cells(x, y + 13).Value 'Radius
    Range("Name").Value = ThisWorkbook.Worksheets(e).Range("Name").Value 'Crane Name
    Range("Tonnage").Value = ThisWorkbook.Worksheets(e).Range("Tonnage").Value 'Tonnage
    Range("Main_Len").Value = ThisWorkbook.Worksheets(e).Range("Main_Len").Value 'Main Boom Length
    Range("CWT").Value = ThisWorkbook.Worksheets(e).Range("CWT").Value 'Counterweight
    Range("Block").Value = ThisWorkbook.Worksheets(e).Range("Block").Value 'Block
    Range("Ball").Value = ThisWorkbook.Worksheets(e).Range("Ball").Value 'Ball
    Range("Rigging").Value = ThisWorkbook.Worksheets(e).Range("Rigging").Value 'Rigging
    Cells(30, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Len").Value 'Main boom angle
    Cells(33, 4).Value = ThisWorkbook.Worksheets(e).Range("Jib_Angle").Value 'Jib length
    ActiveWorkbook.Save
End If

End Sub

