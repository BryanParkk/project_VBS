option explicit
On Error Resume Next

'Declare variables
Dim num1
Dim num2
Dim sum
Dim name, city

const SITE_TITLE = "www.google.com"

'Processing
num1 = 10
num2 = 20

sum = num1 + num2

MsgBox sum

MsgBox "The sum of "& num1 & " and " & num2 " is " & summ, 0, SITE_TITLE

MsgBox "** The sum of "& num1 & "and " & num2 & _
            " is " & sum, 0, SITE_TITLE

