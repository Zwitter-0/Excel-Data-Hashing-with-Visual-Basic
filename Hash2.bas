Attribute VB_Name = "Module1"
Public Function StrHash(text As String) As Long
    Dim i As Long
    StrHash = &H65D5BAAA

    For i = 1 To Len(text)
        StrHash = ((StrHash + AscW(Mid$(text, i, 1))) Mod 69208103) * 31&
    Next
End Function
