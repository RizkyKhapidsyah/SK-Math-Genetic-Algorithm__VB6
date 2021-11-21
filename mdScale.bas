Attribute VB_Name = "mdScale"
Option Explicit
'This module contains the code that will scale the raw value that was output
'by the fitness function by what ever means is required.

Public Function scaleValue(value As Single, target_Value As Single) As Single
'this function is used to scale the raw value of the fitness
'function because the raw value is a number of things:
'-the raw value doesn't approach a max or a min for example
' f(x)=x^2 has a minimum of zero while f(x)=x^3 has an inflection point at
'zero but no minimum or maximum
'-Can be used to scale the raw fitness in such a way as to make
'minor differences between values more magnified for example:
'0.9999987 and 0.9999988 are very close and if all your results are
'this close it is difficult to get a good answer so you could do
'some math manipulation and make the difference more measurable.

'Inorder to solve the equation for a particular value I will scale the
'results accordingly...


'If value < 0 Then
'    scaleValue = value * -100
'    Exit Function
'End If
'Dim absValue As Single ' the absolute value of the value
'Dim absTarget As Single 'absolute value of the target value
'
'absValue = Abs(value)
'absTarget = Abs(target_Value)
'
'If absValue > absTarget Then
'    scaleValue = absValue - absTarget
'    Exit Function
'Else
'    scaleValue = absTarget - absValue 'By scaling the value in this manner
'        'we end up with a non-negative value that approaches zero....
'    Exit Function
'End If
scaleValue = Abs(target_Value - value)


End Function

