Attribute VB_Name = "mdMutation"
Option Explicit
Option Base 1

Function mutation(pop() As chromosome, mrate As Single, range() As varRange) As chromosome()

Dim n As Integer, j As Integer, i As Integer  'index
Dim mutate As Double 'holds a random number from 0 to 1
Dim upperx As Double, upperV As Double
Dim psize As Integer

psize = UBound(pop) 'the number of individuals in the population
upperx = UBound(range) 'the number of DV's
'==================

n = 1   'this loop acomplishes simple mutation ie it deletes the number then replaces with a number between a min and max
'in the future add a mutation that has three levels a small addition or subtraction, a medium and a large one
While n <= psize
    j = 1
    While j <= upperx
        Randomize
        mutate = Rnd
        If mutate < mrate Then
            pop(n).dv(j) = range(j).lower + Rnd * (range(j).upper - range(j).lower)    'randomize the bugger
        End If
        j = j + 1
    Wend
    n = n + 1
Wend

mutation = pop

'this seems to work
End Function

