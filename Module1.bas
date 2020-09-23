Attribute VB_Name = "modPWMeter"
'This password meter is based from the algorithm by Phiras with a few tweaks of my own
'
'You can view the algorithm here
'http://phiras.wordpress.com/2007/04/08/password-strength-meter-a-jquery-plugin/
'
'This is my first submission here in psc
'I have long wanted to submit some code but I can not think of something that is
'unique and has not been submitted before.
'
'Usage: Just add the module in your project then use the function
'       or copy the function to your module.
'
' The function:     x=PWMeter(username, password)
'                   x will have the description of the strength of the password
'                   you can then assign the value of x to a label or something
'
'
'The Password strength procedure is working as the follow:
'We have many cases to care about to know a password strength , so we will present a global variable score , and each case will add some points to score.
'At the end of the algorithm we will decide the password strength according to the score value.
'The cases we have are :
'
'    * If the password matches the username then BadPassword
'    * If the password is less than 4 characters then TooShortPassword
'    * Score += password length * 4
'    * Score -= repeated characters in the password ( 1 char repetition )
'    * Score -= repeated characters in the password ( 2 char repetition )
'    * Score -= repeated characters in the password ( 3 char repetition )
'    * Score -= repeated characters in the password ( 4 char repetition )
'    * If the password has 3 numbers then score += 5
'    * If the password has 2 special characters then score += 5
'    * If the password has upper and lower character then score += 10
'    * If the password has numbers and characters then score += 15
'    * If the password has numbers and special characters then score += 15
'    * If the password has special characters and characters then score += 15
'    * If the password is only characters then score -= 10
'    * If the password is only numbers then score -= 10
'
'    * If score > 100 then score = 100
'
'Now according to score we are going to decide the password strength
'
'    * If 0 < score < 34 then BadPassword
'    * If 34 < score < 68 then GoodPassword
'    * If 68 < score < 100 then StrongPassword

Public Function PWMeter(username As String, pw As String) As String
Dim score As Integer

If username = pw Then PWMeter = "Bad": Exit Function
If InStr(1, pw, username) Then PWMeter = "Bad": Exit Function
If InStr(1, username, pw) Then PWMeter = "Bad": Exit Function
If Len(pw) < 4 Then PWMeter = "TooShort": Exit Function

score = score + Len(pw) * 4
score = score - chkRepetition(1, pw) '(1 char repetition)
score = score - chkRepetition(2, pw) '(2 char repetition)
score = score - chkRepetition(3, pw) '(3 char repetition)
score = score - chkRepetition(4, pw) '(4 char repetition)


Dim IsNumber As Boolean
Dim IsChar As Boolean
Dim IsUpper As Boolean
Dim IsLower As Boolean
Dim IsSymbol As Boolean

Dim CountNumber As Long
Dim CountChar As Long
Dim CountUpper As Long
Dim CountLower As Long
Dim CountSymbol As Long

CountNumber = 0
CountChar = 0
CountUpper = 0
CountLower = 0
CountSymbol = 0

For i = 1 To Len(pw)
    IsNumber = False
    IsChar = False
    IsUpper = False
    IsSymbol = False
    IsLower = False
    IsSymbol = False
    
    If Asc(Mid$(pw, i, 1)) >= 48 And Asc(Mid$(pw, i, 1)) <= 57 Then
        IsNumber = True
        CountNumber = CountNumber + 1
    End If
    If Asc(Mid$(pw, i, 1)) >= 97 And Asc(Mid$(pw, i, 1)) <= 122 Then
        IsLower = True
        CountLower = CountLower + 1
    End If
    If Asc(Mid$(pw, i, 1)) >= 65 And Asc(Mid$(pw, i, 1)) <= 90 Then
        IsUpper = True
        CountUpper = CountUpper + 1
    End If
    If IsLower Or IsUpper Then
        IsChar = True
        CountChar = CountChar + 1
    End If
    If Not (IsNumber Or IsChar) Then
        IsSymbol = True
        CountSymbol = CountSymbol + 1
    End If
Next i
If CountNumber >= 3 Then score = score + 5
If CountSymbol >= 2 Then score = score + 5
If CountLower > 0 And CountUpper > 0 Then score = score + 10
If CountNumber > 0 And CountChar > 0 Then score = score + 15
If CountNumber > 0 And CountSymbol > 0 Then score = score + 15
If CountChar > 0 And CountSymbol > 0 Then score = score + 15
If Len(pw) = CountChar Then score = score - 10
If Len(pw) = CountNumber Then score = score - 10

If score > 100 Then score = 100

Select Case score
Case Is <= 20
    PWMeter = "Very Weak"
Case Is <= 40
    PWMeter = "Weak"
Case Is <= 75
    PWMeter = "Good"
Case Is <= 90
    PWMeter = "Strong"
Case Is > 90
    PWMeter = "Very strong"
End Select
End Function

Private Function chkRepetition(pLen As Long, str As String)
    x = 0
    Dim i As Long, j As Long
    For i = 1 To Len(str)
        For j = i + 1 To Len(str)
            If Mid$(str, i, pLen) = Mid$(str, j, pLen) Then
                x = x + 1
            End If
        Next j
    Next i
    chkRepetition = x
End Function
