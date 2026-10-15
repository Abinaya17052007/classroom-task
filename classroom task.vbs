Option Explicit

'====================================================
' VBScript Client-Side / Windows Script Program
' Variables, Operators, Procedures,
' Conditional Statements, Loops and InputBox/MsgBox
'====================================================

Dim name
Dim age
Dim mark1
Dim mark2
Dim mark3
Dim total
Dim average
Dim i
Dim result
Dim choice


'====================================================
' 1. INPUTBOX - GET USER INPUT
'====================================================

name = InputBox("Enter your name:", "Student Details")

age = InputBox("Enter your age:", "Student Details")

mark1 = InputBox("Enter Mark 1:", "Student Marks")

mark2 = InputBox("Enter Mark 2:", "Student Marks")

mark3 = InputBox("Enter Mark 3:", "Student Marks")


'====================================================
' 2. VARIABLES AND OPERATORS
'====================================================

total = CInt(mark1) + CInt(mark2) + CInt(mark3)

average = total / 3


'====================================================
' 3. MSGBOX - DISPLAY STUDENT DETAILS
'====================================================

MsgBox "Name: " & name & vbCrLf & _
       "Age: " & age & vbCrLf & _
       "Mark 1: " & mark1 & vbCrLf & _
       "Mark 2: " & mark2 & vbCrLf & _
       "Mark 3: " & mark3 & vbCrLf & _
       "Total: " & total & vbCrLf & _
       "Average: " & average, _
       vbInformation, "Student Details"


'====================================================
' 4. CONDITIONAL STATEMENT
'====================================================

If average >= 90 Then

    MsgBox "Grade: A+", vbInformation, "Result"

ElseIf average >= 80 Then

    MsgBox "Grade: A", vbInformation, "Result"

ElseIf average >= 70 Then

    MsgBox "Grade: B", vbInformation, "Result"

ElseIf average >= 60 Then

    MsgBox "Grade: C", vbInformation, "Result"

ElseIf average >= 50 Then

    MsgBox "Grade: D", vbInformation, "Result"

Else

    MsgBox "Result: FAIL", vbCritical, "Result"

End If


'====================================================
' 5. PROCEDURE
'====================================================

Call DisplayMessage()


'====================================================
' 6. FUNCTION
'====================================================

result = AddNumbers(20, 30)

MsgBox "Function Result:" & vbCrLf & _
       "20 + 30 = " & result, _
       vbInformation, "Function"


'====================================================
' 7. FOR LOOP
'====================================================

For i = 1 To 5

    MsgBox "For Loop Value = " & i, _
           vbInformation, "For Loop"

Next


'====================================================
' 8. DO WHILE LOOP
'====================================================

i = 1

Do While i <= 3

    MsgBox "Do While Value = " & i, _
           vbInformation, "Do While Loop"

    i = i + 1

Loop


'====================================================
' 9. SELECT CASE
'====================================================

choice = InputBox( _
    "Enter a number from 1 to 7:" & vbCrLf & _
    "1 - Monday" & vbCrLf & _
    "2 - Tuesday" & vbCrLf & _
    "3 - Wednesday" & vbCrLf & _
    "4 - Thursday" & vbCrLf & _
    "5 - Friday" & vbCrLf & _
    "6 - Saturday" & vbCrLf & _
    "7 - Sunday", _
    "Select Day")


Select Case choice

    Case 1
        MsgBox "Monday", vbInformation, "Day"

    Case 2
        MsgBox "Tuesday", vbInformation, "Day"

    Case 3
        MsgBox "Wednesday", vbInformation, "Day"

    Case 4
        MsgBox "Thursday", vbInformation, "Day"

    Case 5
        MsgBox "Friday", vbInformation, "Day"

    Case 6
        MsgBox "Saturday", vbInformation, "Day"

    Case 7
        MsgBox "Sunday", vbInformation, "Day"

    Case Else
        MsgBox "Invalid Day", vbExclamation, "Error"

End Select


'====================================================
' 10. COOKIE-LIKE PERSISTENT DATA
'====================================================

Dim shell

Set shell = CreateObject("WScript.Shell")

shell.RegWrite _
    "HKCU\Software\VBScriptStudent\Username", _
    name

MsgBox "Student name has been stored successfully." & vbCrLf & _
       "Stored Name: " & name, _
       vbInformation, "Stored Data"

Set shell = Nothing


'====================================================
' END OF PROGRAM
'====================================================

MsgBox "VBScript Program Completed Successfully!", _
       vbInformation, _
       "Completed"


'====================================================
' SUB PROCEDURE
'====================================================

Sub DisplayMessage()

    MsgBox "Welcome " & name & "!" & vbCrLf & _
           "This message is displayed using a Sub Procedure.", _
           vbInformation, _
           "Procedure"

End Sub


'====================================================
' FUNCTION
'====================================================

Function AddNumbers(a, b)

    AddNumbers = a + b

End Function