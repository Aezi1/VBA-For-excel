Attribute VB_Name = "Module1"
' This VBA is for RED text to red color and yellow to yellow and green to green, don't worry about turn rediscover to "red"discover, I have changed the codes helped by chatgpt.
Sub ChangeColorOfSpecificWordsOnly()
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Define the colors with a darker yellow
    Dim redColor As Long, yellowColor As Long, greenColor As Long
    redColor = RGB(255, 0, 0)
    yellowColor = RGB(255, 215, 0) ' A darker shade of yellow
    greenColor = RGB(0, 128, 0)

    ' Loop through each cell in the active sheet
    For Each cell In ws.UsedRange.Cells
        Call ChangeWordColorInCell(cell, " red ", redColor)
        Call ChangeWordColorInCell(cell, " yellow ", yellowColor)
        Call ChangeWordColorInCell(cell, " green ", greenColor)
    Next cell
End Sub

Sub ChangeWordColorInCell(cell As Range, searchText As String, textColor As Long)
    Dim pos As Integer
    Dim textLen As Integer
    Dim cellText As String
    cellText = " " & cell.Text & " "
    searchText = " " & searchText & " "
    textLen = Len(searchText)
    
    pos = InStr(1, cellText, searchText, vbTextCompare)
    
    ' Loop through all occurrences of the search text and change the color
    While pos > 0
        cell.Characters(Start:=pos, Length:=textLen).Font.Color = textColor
        pos = InStr(pos + textLen, cellText, searchText, vbTextCompare)
    Wend
End Sub

