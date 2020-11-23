
Public Sub Main()

    Dim sMsg As String

    sMsg = GetFonts(ActiveDocument)
    MsgBox "The fonts in this document are:" & vbNewLine & vbNewLine & sMsg
End Sub

Private Function GetFonts(ByVal oDocument As Document) As String

    Dim oParagraph      As Paragraph
    Dim i               As Integer
    Dim oWord           As Words
    Dim sFontType       As String
    Dim sFontSize       As String
    Dim sMsg            As String

    For Each oParagraph In oDocument.Paragraphs
        For i = 1 To oParagraph.Range.Characters.Count
            sFontSize = ""
            sFontType = oParagraph.Range.Characters(i).Font.Name
            sFontSize = CStr(oParagraph.Range.Characters(i).Font.Size)
            If InStr(1, sMsg, sFontType) = 0 Then
                sMsg = sMsg & sFontType & " " & sFontSize & vbNewLine
            End If
        Next
    Next
    GetFonts = sMsg
End Function
