Attribute VB_Name = "mdFitness"
Option Explicit
Option Base 1


'in order to be able to solve an equation x + y + z + 15 = 28
'The fitness is easyily calculated, the (optimal value - the ga value)/(optimal value) and take the absolute value of the result, the closest one to zero is the optimal solution


'Remember if you decide to use a different type of function,
'one that has a naturally max or min value, then be sure to change the
'scaling in the mdScale mod

Public Function fitness(var() As Single)
'var - an array holding the values to input in to the fitness function

'since we are only going to deal with an equation that has 4 variables
'there is no need to determine the upper and lower bounds of the var array

fitness = Round(var(1) + var(2) + var(3) + 1, 2)
            'I used the round function to get rid of the decimal error
            'If the round function is not used, then the GA will
            'most likely never get to the value of 28.....Because of rounding errors
            '

End Function

