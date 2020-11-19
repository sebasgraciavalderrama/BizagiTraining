VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ReplaceFont()
    Dim obSingleWord As Range
    Dim obSingleWord1 As Range
    Dim obSingleWord2 As Range
    Dim objDoc As Document
    Dim BodyText As String
    Dim HeadingOne As String
    Dim HeadingTwo As String

    'Font name goes here...'
    BodyText = "Segoe UI Semilight"
    HeadingOne = "Segoe UI Semibold"
    HeadingTwo = "Segoe UI Semibold"

    Set objDoc = ActiveDocument

    With objDoc
        For Each objSingleWord In .Words 'For body text'
            If objSingleWord.Font.Name = BodyText Then
                objSingleWord.Font.Name = "Open Sans Light (Body)"
                objSingleWord.Font.Size = 11
                objSingleWord.Font.TextColor.RGB = RGB(65, 86, 104)
            End If
        Next objSingleWord
        
        For Each objSingleWord1 In .Words 'For Headings'
            If objSingleWord1.Font.Name = HeadingOne Then
                objSingleWord1.Font.Name = "Open Sans ExtraBold"
                If objSingleWord1.Font.Size < 18 Then 'For Subtitles'
                    objSingleWord1.Font.Size = 13
                    objSingleWord1.Font.TextColor.RGB = RGB(240, 139, 29) 'Orange font'
                ElseIf objSingleWord1.Font.Size >= 18 Then 'For titles'
                    objSingleWord1.Font.Size = 21
                    objSingleWord1.Font.Bold = True
                    objSingleWord1.Font.TextColor.RGB = RGB(65, 86, 104)
                End If
            End If
        Next objSingleWord1
    End With
End Sub

