Attribute VB_Name = "modHash"
Option Explicit

    
    'hash_generator Start----->
Public Function generate_Hash(ByVal mytext As String) As String
    'calling funtions'
    generate_Hash = fun_hash(mytext)
End Function

    'end of hash_genrator---->


Public Function fun_hash(mytext As String, Optional bB64 As Boolean = 0) As String
    '----------< SHA256 () >----------
    
    '< setup_Encoders >
    Dim Encoder As Object
    Set Encoder = CreateObject("System.Text.UTF8Encoding")
    
    Dim Encoder_fun_hash As Object
    Set Encoder_fun_hash = CreateObject("System.Security.Cryptography.SHA256Managed")
    '</ setup_Encoders >
    
    '< encode >
    Dim TextToHash() As Byte
    TextToHash = Encoder.GetBytes_4(mytext)
    '</ encode >
    
    
    '*create Byte Arrays
    Dim bytes() As Byte
    bytes = Encoder_fun_hash.ComputeHash_2((TextToHash))
    
    '< convert and return >
    If bB64 = True Then
        
       fun_hash = ConvToBase64String(bytes)
    Else
       fun_hash = ConvToHexString(bytes)
    End If
    '</ convert and return >
    
    '< close >
    Set Encoder = Nothing
    Set Encoder_fun_hash = Nothing
    '</ close >
    '----------</ SHA256 () >----------
End Function




Public Function ConvToBase64String(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.text, vbLf, "")
    
    Set oD = Nothing

End Function

Public Function ConvToHexString(vIn As Variant) As Variant

    Dim oD As Object
      
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.text, vbLf, "")
    
    Set oD = Nothing

End Function

