
Sub Update_Collection() 'Button: Update List

'make sheet temporarily unprotected
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Commemorative Coins")
    
'no password
ws.Unprotect Password:=""


'Color_Commemorative Coins()

Sheets("Commemorative Coins").Activate
Range("E$4").Select

'change color depending on whether you own that coin (red = not owned; green = owned)
If ActiveCell.Value = "0" Then
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Interior.Color = RGB(255, 51, 0)
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Interior.Color = RGB(255, 51, 0)
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Interior.Color = RGB(255, 51, 0)
Else
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Interior.Color = RGB(146, 208, 80)
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Interior.Color = RGB(146, 208, 80)
    ActiveCell.Offset(0, -1).Select
    ActiveCell.Interior.Color = RGB(146, 208, 80)
End If

Do While ActiveCell.Offset(1, 3).Value = "0" Or ActiveCell.Offset(1, 3).Value > 0
    ActiveCell.Offset(1, 3).Select

    If ActiveCell.Value = "0" Then
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Interior.Color = RGB(255, 51, 0)
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Interior.Color = RGB(255, 51, 0)
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Interior.Color = RGB(255, 51, 0)
    Else
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Interior.Color = RGB(146, 208, 80)
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Interior.Color = RGB(146, 208, 80)
        ActiveCell.Offset(0, -1).Select
        ActiveCell.Interior.Color = RGB(146, 208, 80)
    End If
Loop


'Color_Mintmarks()
    
    Sheets("Commemorative Coins").Activate
    Range("B4").Select
    
    Do Until ActiveCell.Value = "Germany"
        ActiveCell.Offset(1, 0).Select
    Loop
    
    ActiveCell.Offset(0, 5).Select
    
    Do While ActiveCell.Value <> ""
    
        Do While ActiveCell.Value <> ""
    
            If ActiveCell.Value = "0" Then
                ActiveCell.Interior.Color = RGB(255, 51, 0)
            Else
                ActiveCell.Interior.Color = RGB(146, 208, 80)
            
            End If
        
        ActiveCell.Offset(0, 1).Select
        
        Loop
    
    ActiveCell.Offset(1, -5).Select
    
    Loop

'Update_Count()

    Sheets("Commemorative Coins").Activate
    
    'declaring i as the number of collected coins and total as the number of total coins
    Dim i As Integer
    Dim total As Integer
    
    'counting the number of total coins
    Range("E4").Select
    total = 0
    
    Do While ActiveCell.Value <> ""
        total = total + 1
        ActiveCell.Offset(1, 0).Select
    Loop
    
    'counting the number of collected coins
    Range("E4").Select
    i = 0
    
    Do While ActiveCell.Value <> ""
        If ActiveCell.Value > 0 Then
            i = i + 1
        End If
        
        ActiveCell.Offset(1, 0).Select
        
    Loop
    
    
    Range("H9").Value = i & " of " & total
    
    'reactive protection
        ws.Protect Password:="", UserInterfaceOnly:=True

End Sub

Sub ListDuplicates() 'Button: Update Duplicates

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim i As Long
    Dim countOwned As Long
    Dim duplicates As Long
    Dim rng As Range

    ' Source: "Commemorative Coins"
    Set wsSource = Worksheets("Commemorative Coins")
    
    'make sheet temporarily unproteced
    wsSource.Unprotect Password:=""
    
    ' Target: "Duplicate" – create or empty
    On Error Resume Next
    Set wsTarget = Worksheets("Duplicates")
    If wsTarget Is Nothing Then
        Set wsTarget = Worksheets.Add
        wsTarget.Name = "Duplicates"
    Else
        wsTarget.UsedRange.ClearContents ' only delete content, keep form
    End If
    On Error GoTo 0

    ' add headings
    wsTarget.Range("A1").Value = "Country"
    wsTarget.Range("B1").Value = "Year"
    wsTarget.Range("C1").Value = "Title"
    wsTarget.Range("D1").Value = "Duplicates"

    ' determine last row in Source
    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    targetRow = 2 ' start in row 2 in Source-sheet

    ' iterative through each coin
    For i = 4 To lastRow
        countOwned = wsSource.Cells(i, "E").Value
        If IsNumeric(countOwned) And countOwned >= 2 Then
            duplicates = countOwned - 1
            wsTarget.Cells(targetRow, "A").Value = wsSource.Cells(i, "B").Value ' Country
            wsTarget.Cells(targetRow, "B").Value = wsSource.Cells(i, "C").Value ' Year
            wsTarget.Cells(targetRow, "C").Value = wsSource.Cells(i, "D").Value ' Title
            wsTarget.Cells(targetRow, "D").Value = duplicates ' Duplikate
            targetRow = targetRow + 1
        End If
    Next i

    ' set borders for all newly added coins
    If targetRow > 2 Then
        Set rng = wsTarget.Range("A1:D" & targetRow - 1)
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If

    MsgBox "Duplicates were updated", vbInformation
    
    'reactivate protection
    wsSource.Protect Password:="", UserInterfaceOnly:=True
    wsTarget.Protect Password:="", UserInterfaceOnly:=True
    
End Sub

