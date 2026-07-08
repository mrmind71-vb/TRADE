Attribute VB_Name = "SerializeMdl"
Public Function Serialize(documentStructure As String) As String
    Dim json As New ChilkatJsonObject
    Dim sb As New ChilkatStringBuilder
    
    success = json.Load(documentStructure)
    If json.Size = 0 Then
        Serialize = """" & documentStructure & """"
        Exit Function
    End If
                    
    For i = 0 To json.Size - 1
        If json.TypeAt(i) <> 4 Then
            sb.Append """" & UCase(json.NameAt(i)) & """"
            sb.Append Serialize(json.StringAt(i))
        Else
            sb.Append """" & UCase(json.NameAt(i)) & """"
            
            For i2 = 0 To json.ArrayAt(i).Size - 1
                sb.Append """" & UCase(json.NameAt(i)) & """"
                sb.Append Serialize(json.ArrayAt(i).StringAt(i2))
            Next
        End If
    Next
    Serialize = sb.GetAsString()
End Function
Public Function Serialize2(jsonStr As String) As String
Dim json As New ChilkatJsonObject
success = json.Load(jsonStr)

If success = 0 Then Exit Function

If json.Size = 0 Then
    Serialize2 = myFix(jsonStr)
    Exit Function
End If

For i = 0 To json.Size - 1
    If json.TypeAt(i) = 1 Or json.TypeAt(i) = 2 Or json.TypeAt(i) = 5 Or json.TypeAt(i) = 6 Then
        Serialize2 = Serialize2 & UCase(myFix(json.NameAt(i)))
        Serialize2 = Serialize2 & myFix(json.StringAt(i))
    ElseIf json.TypeAt(i) = 4 Then
        Serialize2 = Serialize2 & UCase(myFix(json.NameAt(i)))
        For i2 = 0 To json.ArrayAt(i).Size - 1
            Serialize2 = Serialize2 & UCase(myFix(json.NameAt(i)))
            Serialize2 = Serialize2 & Serialize2(json.ArrayAt(i).StringAt(i2))
        Next
    ElseIf json.TypeAt(i) = 3 Then
        If json.NameAt(i) <> "" Then
            Serialize2 = Serialize2 & UCase(myFix(json.NameAt(i)))
        End If
        Serialize2 = Serialize2 & Serialize2(json.ObjectAt(i).Emit())
    End If
Next
End Function
Private Function myFix(pString As String) As String
myFix = """" & pString & """"
End Function


