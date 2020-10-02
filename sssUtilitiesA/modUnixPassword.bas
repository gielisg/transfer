Attribute VB_Name = "modUnixPassword"
Option Explicit

Function DecryptInformix(ByVal UserId As String, ByVal Password As String) As String
    Dim UserIdChar As String
    Dim UserIdByte As Integer
    Dim PasswordChars As String
    Dim PasswordByte As Integer
    Dim PasswordIndex As Integer
    Dim Result As Integer
    Dim ResultStr As String
    
    DecryptInformix = ""

    If Left(Password, 2) <> "EP" Then
        MsgBox "Invalid Password.", vbCritical + vbOKOnly, "Invalid Password."
        Exit Function
    End If
  
    Password = Mid(Trim(Password), 3)
    PasswordIndex = 0
 
    While Password <> ""
        PasswordIndex = PasswordIndex + 1
        PasswordChars = Left(Password, 3)

        If (Not IsNumeric(PasswordChars)) Then
            MsgBox "Not numeric: " & PasswordChars, vbCritical + vbOKOnly, "Invalid Password."
            Exit Function
        End If

        PasswordByte = PasswordChars
    
        If (PasswordIndex > Len(UserId)) Then
            UserIdByte = 0
        Else
            UserIdChar = Mid(UserId, PasswordIndex, 1)
            UserIdByte = Asc(UserIdChar)
        End If

        Result = PasswordByte Xor UserIdByte
        If Result > 0 Then
            ResultStr = ResultStr & Chr(Result)
        End If
        Password = Mid(Password, 4)
    Wend

    DecryptInformix = ResultStr

End Function

Function EncryptInformix(ByVal UserId As String, ByVal Password As String) As String
    Dim UserIdChar As String
    Dim UserIdByte As Integer
    Dim PasswordChar As String
    Dim PasswordByte As Integer
    Dim PasswordIndex As Integer
    Dim Result As Integer
    Dim ResultStr As String
    
    ResultStr = "EP"
    
    For PasswordIndex = 1 To 18

        If (PasswordIndex > Len(Password)) Then
            PasswordByte = 0
        Else
            PasswordChar = Mid(Password, PasswordIndex, 1)
            PasswordByte = Asc(PasswordChar)
        End If
      
        If (PasswordIndex > Len(UserId)) Then
            UserIdByte = 0
        Else
            UserIdChar = Mid(UserId, PasswordIndex, 1)
            UserIdByte = Asc(UserIdChar)
        End If
       
        Result = PasswordByte Xor UserIdByte
    
        If (Result < 10) Then
            ResultStr = ResultStr & "  " & Result
        ElseIf (Result < 100) Then
            ResultStr = ResultStr & " " & Result
        Else
            ResultStr = ResultStr & "" & Result
        End If
    Next

    EncryptInformix = ResultStr
End Function

