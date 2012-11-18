Attribute VB_Name = "CellWalker"
Sub CellWalker()

    Dim regexPattern As String
    regexPattern = InputBox("Input a Regex-Pattern like the Perl", "Regex-Pattern")
    If regexPattern = "" Then
        Exit Sub
    End If

    Dim aRegex As Object
    Set aRegex = CreateObject("VBScript.RegExp")
        aRegex.Pattern = regexPattern
        aRegex.Global = True
        aRegex.IgnoreCase = False
    
    Dim colorIndexInput As String
    Dim colorIndex As Integer
    colorIndexInput = InputBox("Input a color index. ex) 3=Red, 4=Green, 5=Blue, 6=Yellow ..", "Color index")
    If Not IsNumeric(colorIndexInput) Then
        Exit Sub
    End If
    colorIndex = Val(colorIndexInput)

    Dim rangeKey As String
    Dim ranges As Object
    rangeKey = InputBox("Input a range key. ex) 'A1:Z100', ''(auto)", "Range")
    If rangeKey = "" Then
        Set ranges = ActiveSheet.UsedRange
    Else
        Set ranges = ActiveSheet.Range(rangeKey)
    End If

    Dim match As Object
    Dim matches As Object
    For Each aRange In ranges
        If Not IsError(aRange.Value) Then ' Ignore #VALUE!, #DIV/0!
            aRange.Select
            Set matches = aRegex.Execute(aRange.Value)
            For Each match In matches
                With ActiveCell.Characters(Start:=match.FirstIndex + 1, Length:=match.Length).Font
                    .colorIndex = colorIndex
                End With
            Next
        End If
    Next
End Sub
