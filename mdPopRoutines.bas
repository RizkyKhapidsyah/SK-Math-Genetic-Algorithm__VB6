Attribute VB_Name = "mdPopRoutines"
Option Explicit
Option Base 1

'This module contains the population routines ie, the routines
'that are used to generate a starting population...
'===========================================


Public Function popBuilder(numInd As Integer, numVar As Integer, rng() As varRange) As chromosome()
'numInd - is the number of individual chromosomes within the population
'numVar - the number of variables that the chromosome will have
'           eg. if f(x,y) = x+y then the number of variables would
'           be 2...
'rng - an array of varRange that is used to hold the upper and lower bounds
'       of each of the dv values


'it returns an array of chromosomes....
Dim i As Integer 'index
Dim j As Integer 'index
Dim tempPop() As chromosome
ReDim tempPop(1 To numInd) 'set the size of the chromosome array

For i = 1 To numInd  'this is the loop that creates the individual chromosomes
    ReDim tempPop(i).dv(1 To numVar) 'also need to make sure that dv is set to the appropriate size
    
    For j = 1 To numVar
        Randomize
        tempPop(i).dv(j) = rng(j).lower + Rnd * (rng(j).upper - rng(j).lower) 'by creating the random numbers in this fashion it ensures that the numbers will be valid and within the correct range
    Next j
Next i

popBuilder = tempPop 'return the newly created population
End Function

