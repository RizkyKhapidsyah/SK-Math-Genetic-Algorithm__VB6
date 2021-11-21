Attribute VB_Name = "mdDefn"
Option Explicit
'========================================
Public Type chromosome
    dv() As Single
    value As Single
    fitness As Single
End Type

'The above type is similar to a C struct.
'it is used: dim x as chromosome 'and defines x to have the properties of dv(), value and fitness
'The reason a type is used is to eliminate some of the loops
'that would be necessary to do some of the calculations....

'chromosome methods defined:
'dv - is short for decision variable or if you will a variable
'   these are the values that the GA will try to optimize. for example
'   if your fitness function were f(x)=x^2 then you would only have one dv
'   and it would be x. Now the reason that dv is an array is to make this
'   GA as flexible as possible.
'value - This is the result of evaluating the dvs in the fitness function.

'fitness - In most cases this is equal to the value. You may ask yourself
'           why included this value. Well, if you wish to scale or normalize
'           the actual value produced by the fitness function to try and achieve better results.
'           Also it can be used to easily change your goal (maximization or minimization) of the GA
'           without any significant changes to the algorithm
'           Note: This algorithm will only work with the fitness value, the "value" variable will simply
'           be used to store the raw unscaled output of the fitness function

'A further aside, this use of a type will become clear as you see it used in actual code
'Also if you need more precision simply change the dv type to double....
'===================================

Public Type varRange
    lower As Single
    upper As Single
End Type

'varRange is used to hold the upper and the lower bounds of a particular
'dv
'=======================
