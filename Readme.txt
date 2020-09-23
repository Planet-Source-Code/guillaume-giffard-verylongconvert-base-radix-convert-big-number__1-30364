'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VeryLongConvert : function that converts a huge number as string from a base to
'                  another one
'
'                  Version : 1.01
'                  Author:  Guillaume GIFFARD
'                  Date : 01/03/2002
'                  Mail : Guiland@mail.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUTS :  * Word As String : the huge number, up to 32,000 digits, to convert
'
'          * FromBase As Integer : the base in witch Word is written
'
'          * ToBase As Integer : the base in witch Word is to convert
'
'          * Separator As String : this Optional variable is the decimal separator,
'          usely the point and sometimes the comma
'
'          FromBase and ToBase are integers from 2 to 36
'
'OUTPUTS : * the function returns the huge number converted from FromBase to
'          ToBase as string. It returns "" if Word is empty or if FromBase or
'          ToBase is not between 2 and 36
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Here is the public code

Public Const B_BIN As Integer = 2
Public Const B_OCT As Integer = 8
Public Const B_DEC As Integer = 10
Public Const B_HEX As Integer = 16

Public Const DEFAULT_SEPARATOR As String = "."
Public Const COMMA_SEPARATOR As String = ","

Public Function VeryLongConvert(Word As String, FromBase As Integer, ToBase As Integer, Optional Separator As String = DEFAULT_SEPARATOR) As String

'Example :

'This example illustrates the VeryLongConvert function. To try this example, paste the
'code into the Declarations section of a form that contains two TextBox controls and a
'CommandButton control, and then press F5, enter a number in Text1 and click on
'Command1. It will convert it from decimal to hexadecimal. You can convert from any
'base to another by replacing 10 and 16 by the bases you want.

Private Sub Command1_Click()
    Text2.Text = VeryLongConvert(Text1.Text, 10, 16)
End Sub

'How it works :

'I won't go in the details but it works with divisions to convert a number from a base to another. e.g. : if I want to convert 247 from decimal base to binary base I'll do that way :
'	numerator|denominator		LSB : Less significant bit
'	remainder/quotient		MSB : Most significant bit

'	247|2
'LSB	  1/123|2
'	      1/61|2
'		 1/30|2
'		    0/15|2
'		       1/7|2
'			 1/3|2
'			   1/1|2
'MSB			     1/0

'	So, 247 in decimal is written 11110111 in binary















'Try my example with FrmTest.
'Give your comments at Guiland@mail.com
'Thank you for using one of my programmes !