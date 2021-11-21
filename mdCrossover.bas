Attribute VB_Name = "mdCrossover"
Option Explicit
Option Base 1


'needs to modified
Function xover(pop() As chromosome, srows() As Integer, range() As varRange) As chromosome()
'pop is the population to be crossed over vector,
'srows - an array containing the chromosomes that are to undergo crossover
'       The structure of the array is such that the first element contains the
'       index of the chromosome within the pop array that is to be used in the
'       crossover. The srows array contains the same number of elements as the pop
'       array but, not all the elements have choosen individuals....
'       This is what the array could look like (1,1,2,3,4,3,5,0,0,0,0...) The zeros indicate
'       that there are nomore elements to be selected. When one of these zero
'       values is reached, that means there are no more chromosomes to be crossed over.


'====================
Dim n As Integer, j As Integer  ' index
Dim len_srows As Integer, xvrPoint As Integer
Dim beta As Double ' used to hold the random number between 0 and 1
Dim x As Double, y As Double ' used to hold the values to be crossed over
Dim temp1 As Double, temp2 As Double
Dim psize As Integer
Dim numDVs As Integer

len_srows = UBound(srows) ' length of the array
numDVs = UBound(range) 'the number of variables that the GA is using
temp2 = UBound(pop)

n = 1
While n < len_srows ' the loop that does the swapping
    Randomize
    xvrPoint = 1 + Int(Rnd * (numDVs - 1)) ' the crossover point
    Randomize
    beta = Rnd  ' this is a random number used in the aritmetic operator
    
    j = xvrPoint
    'perform arithmetic operation first, then swap the rest of the DV's
    
    If srows(n) = 0 Then GoTo exitLoop 'make sure that srows() still has some chromosomes to crossover
    
    x = pop(srows(n)).dv(j)   'from parent 1
    y = pop(srows(n + 1)).dv(j) 'from parent 2
        
    pop(srows(n)).dv(j) = beta * y + (1 - beta) * x      'finalx1
    pop(srows(n + 1)).dv(j) = beta * x + (1 - beta) * y ' finalx2
    j = j + 1
        
    'now swap the rest of the DV's
    While j <= numDVs
        temp1 = pop(srows(n)).dv(j)   'next DV from parent 1
        temp2 = pop(srows(n + 1)).dv(j) 'next DV from parent 2
        
        pop(srows(n)).dv(j) = temp2
        pop(srows(n + 1)).dv(j) = temp1
        
        j = j + 1
    Wend
    
    
    n = n + 2
Wend
exitLoop:

xover = pop

End Function

