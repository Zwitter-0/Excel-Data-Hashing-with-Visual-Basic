Attribute VB_Name = "Module2"
Public Function StrHash64(text As String) As String
    Dim i&, h1&, h2&, c&
    h1 = &H65D5BAAA
    h2 = &H2454A5ED

    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        h1 = ((h1 + c) Mod 69208103) * 31&
        h2 = ((h2 + c) Mod 65009701) * 33&
    Next

    StrHash64 = Right("00000000" & Hex(h1), 8) & Right("00000000" & Hex(h2), 8)
End Function
