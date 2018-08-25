Attribute VB_Name = "Module1"
'-----------------Created by Nicholas Struckmeyer 8/10/2018---------------------------

'Method for exporting values as CSV from a workbook by reading cell names based on
' input from an "Instructions" tab
Sub CreateCSV()
    'Final results array to send to CSV
    Dim exportCol As Collection
    Dim instructions As Collection
    'Instructions from "Instructions" tab
    Set instructions = getInstructions()
    If instructions Is Nothing Then 'Don't do nothin' if there ain't nothin'
        Exit Sub
    End If
    
    Set exportCol = getValuesToExport(instructions)
    
    If exportCol Is Nothing Then
        MsgBox "Export values not retrieved, check export collection."
        Exit Sub
    End If
    
    Dim csvPath As String: csvPath = getFilePath()
    
    If (csvPath = Null Or csvPath = vbNullString) Then
        MsgBox "No File Selected"
        Exit Sub
    End If
    
    'Exports the values to a CSV file (TODO create file picker)
    Close #1
    Open csvPath For Output As #1
    For i = 1 To exportCol.count
        Print #1, exportCol.Item(i)
    Next i
    Close #1


End Sub

'Retrieves a Collection of Strings given instructions in the form of a Collection
' of String Arrays
Function getValuesToExport(instructions As Collection) As Collection
    Dim wbk As Workbook: Set wbk = ThisWorkbook
    Dim exportCol As Collection: Set exportCol = New Collection
    Dim instructionCount As Integer: instructionCount = instructions.count
    Dim outerI As Integer
    
    'Loop through instruction sets (columns representing tabs in workbook)
    For outerI = 1 To instructionCount
        Dim instArray() As String: instArray = instructions.Item(outerI)
        Dim arrCount As Integer: arrCount = UBound(instArray, 1) - LBound(instArray, 1)
        Dim sheetName As String: sheetName = instArray(0)
        Dim selectedSheet As Worksheet: Set selectedSheet = wbk.Sheets(sheetName)
        Dim i As Integer
        
        'Note: must start on element 1 becuase the 0 element is the tab name
        For i = 1 To arrCount - 1
           exportCol.Add (selectedSheet.Range(instArray(i)).Value)
           
        Next i
    Next outerI
    
    Set getValuesToExport = exportCol

End Function

'Retrieves a Collection of String Arrays representing instructions on how to
' retrieve the data from each tab (a list of cell names on each tab)
Function getInstructions() As Collection
    Dim wbk As Workbook: Set wbk = ThisWorkbook
    'Instructions Sheet
    Dim instSheet As Worksheet
    'If no Instructions tab exists in workbook exit the function
    On Error GoTo NoTabErr:
    Set instSheet = ThisWorkbook.Sheets("Instructions")

    'Collection of Arrays representing instructions from each column from Instructions tab
    Dim arrCollection As Collection: Set arrCollection = New Collection
    'Whether there is another column with instructions
    Dim nextColumn As Boolean: nextColumn = True
    Dim startColumn As Integer: startColumn = 1
    
    Dim startRange As Range: Set startRange = instSheet.Cells(1, startColumn)
    
    'Check for new instructions exist
    If startRange.Value = "" Then
    'Exit Function
    nextColumn = False
    End If
    
    While nextColumn = True
        Dim instArray() As String
        Dim cell As Range
        Dim cellCount As Integer: cellCount = 0
        
        'Select instructions
        instSheet.Activate
        Dim instructRange As Range
        
        'Set the instruction range. If only one instruction, DON'T Ctrl + Shift + Down
        If Cells(2, startColumn).Value = "" Then
            Set instructRange = Cells(1, startColumn)
        Else
            Set instructRange = Range(Cells(1, startColumn), _
                Cells(1, startColumn).End(xlDown))
        End If
        
        'Create instruction array and add to collection
        ReDim instArray(0 To instructRange.count)
        For Each cell In instructRange.Cells
            instArray(cellCount) = cell.Value
            'cell.Interior.ColorIndex = cellCount + 10
            cellCount = cellCount + 1
        Next cell
        arrCollection.Add (instArray)
    
        'Go to next column
        startColumn = startColumn + 1
        Set startRange = instSheet.Cells(1, startColumn)
        
        'Check for new instructions exist
        If startRange.Value = "" Then
        nextColumn = False
        End If
    Wend
    
    Set getInstructions = arrCollection
    Exit Function
    
    
NoTabErr:
    MsgBox "Please create Instructions tab"
    Exit Function

End Function

'Retrieves the path of file to select
Function getFilePath() As String
    Dim varResult As Variant
    'displays the save file dialog
    varResult = Application.GetSaveAsFilename( _
        FileFilter:="CSV (Comma delimited)(*.csv),*.csv,")
    'checks to make sure the user hasn't canceled the dialog
    If varResult <> False Then
        getFilePath = varResult
        
    Else
        
    End If
    
    
End Function
