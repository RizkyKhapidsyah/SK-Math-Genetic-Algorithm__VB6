Attribute VB_Name = "mdSelection"
Option Explicit
Option Base 1

'There are two main types of selection mechanisms, tournament and roulette.
'There are also a number of variants based on both of these but I am going
'to employ tournament selection as it seems to be the most effective and
'the easiest to code...

'the tournament selection strategy selects a number of individuals
'for example four, the fittest of the four is inserted into the
'new population. This process is repeated untill the population is
'full

'For now the tournament function creates a new population out of the old one
'replacing the old population with the winners of the tournaments.....
'This function can be modifed to replace a certain number of chromosomes
' choosen at random, or the weakest members of the population etc...
'for simplicities sake I have simply replaced the entire population
Public Function tournament(pop() As chromosome, participants As Integer) As chromosome()
'pop - the population of chromosomes
'participants - the number of chromosomes that will compete with each other at
'               any one time....

Dim popSize As Integer, i As Integer, k As Integer

Dim selected() As Integer
ReDim selected(1 To participants)

popSize = UBound(pop)
Dim newPop() As chromosome
ReDim newPop(1 To popSize)

Dim temp As chromosome 'used to hold the best chromosome for each tournament

i = 1
While i <= popSize
    'generate random numbers representing the individuals to be selected for
    'tournament
    For k = 1 To participants
        selected(k) = 1 + Rnd * (popSize - 1)
    Next k
    'find the best out of the tournament participants
    temp = pop(selected(1)) ' assign the temporary to the first selected
    For k = 2 To participants
        If temp.fitness < pop(k).fitness Then
            temp = pop(k)
        End If
    Next k
    'insert the winner into the new population
    newPop(i) = temp
        
    i = i + 1
Wend


tournament = newPop

End Function

'=====================================

'the code below performs roulette wheel selection, basically a
'chromosome is assigned a sector of a certain size based on it's
'fitness value

Public Function roulette(popVector() As chromosome) As chromosome()
    Dim fitness As Double 'total fitness of the population
    Dim n As Integer, popSize As Integer  ' n is an index, and popSize is a variable to hold the upperbound of the population array
    Dim selection_prob() As Double 'probability of an individuals selection
    Dim cumm_prob() As Double ' the cummulative probability of each indivdual
    Dim newPopulation() As chromosome ' the new population after roulette selection
    Dim selectedPop() As Integer  'is an array that holds the positions of the selected individuals to be sent into the roulette wheel selection process
       
    popSize = UBound(popVector) 'upper bounds
    
    'determine the total fitness of the population
    n = 1
    fitness = 0 ' initialize the value to zero
    While n <= popSize
        fitness = fitness + popVector(n).fitness
        n = n + 1
    Wend
    
    selection_prob = selprob(popVector, fitness, popSize)
    cumm_prob = cummprob(selection_prob, popSize)
       
    
    ReDim selectedPop(1 To popSize)
    ReDim newPopulation(1 To popSize)
    
    selectedPop = selectedrows(cumm_prob, popSize)
    newPopulation = new_pop(popVector, selectedPop, popSize) 'returns the new population
    
    roulette = newPopulation 'returns the new population to the main program
    
      

    
End Function

'selprob goes with the roulette wheel function
Function selprob(eval() As chromosome, fit As Double, pops As Integer) As Double()
Dim temp() As Double
Dim n As Integer
ReDim temp(1 To pops)

n = 1
While n <= pops
    temp(n) = eval(n).fitness / fit
    n = n + 1
Wend
selprob = temp

End Function

Function cummprob(selprob() As Double, psize As Integer) As Double()
Dim n As Integer, i As Integer, temp As Double
Dim temp_sto() As Double
n = 1
temp = 0
ReDim temp_sto(1 To psize)

While n <= psize
    i = 0
    Do
        i = i + 1
        temp = temp + selprob(i)
    Loop Until i = n
    temp_sto(n) = temp
    temp = 0
    n = n + 1
Wend
If temp_sto(psize) = 1 Then 'checks to see if the alst value is one
Else                        'if it isn't then it is changed
   temp_sto(psize) = 1
End If

cummprob = temp_sto

End Function


Function selectedrows(cprob() As Double, psize As Integer) As Integer()
Dim rowselected() As Integer
Dim temp As Double
Dim n As Integer, p As Integer
n = 1
ReDim rowselected(1 To psize)

While n <= psize
    Randomize
    temp = Rnd
    p = 1
    Do While p <= psize
        If cprob(p) >= temp Then
            rowselected(n) = p
            Exit Do
        End If
        If p = psize Then ' if nothing is selected then the last one automatically gets picked
            rowselected(n) = p
        End If
        
        p = p + 1
    Loop

    n = n + 1
Wend

selectedrows = rowselected
End Function

Function new_pop(pop() As chromosome, selpop() As Integer, p_size As Integer) As chromosome()
Dim r As Integer
Dim pop_new() As chromosome
ReDim pop_new(1 To p_size)

r = 1
While r <= p_size
    pop_new(r) = pop(selpop(r)) 'puts the selected chromosomes into the new population that will undergo crossover
    r = r + 1
Wend

new_pop = pop_new

End Function
'======end of roulette wheel selection==================
'[[[[[[[[[[[[[[[[[[[]]]]]]]]]]]]]]]]]]]]]]

Function pick_mates(xvr As Single, p_size As Integer) As Integer() 'this functions returns an array of values to be crossed over, it also contains nulls which are not used
'xvr is the crossover probability
'p_size is the population size

'The purpose of this function is to select individual chromosomes from
'the population by using the crossover probability

'It returns an integer array containing the indexes of the

Dim n As Integer 'index
Dim q As Integer 'index
Dim numofsel As Integer 'holds the number of individuals selected inorder to determine wheather there is an even or odd number
Dim temp As Double 'holds a random number from 0 to 1
Dim choice As Integer ' choice is used to go through the population and delete a row at random
Dim selcross() As Integer ' the position of the individuals choosen for mating
Dim temp1() As Integer ' an intermediat variable
Dim ranrow As Integer ' holds a random row that is to be deleted
Dim done As Boolean 'wheather the algorithm is finished adding or deleting a row

ReDim selcross(1 To p_size)
ReDim temp1(1 To p_size)

n = 1
While n <= p_size
    selcross(n) = 0 ' initializing selcross to zero, this will add in deleting rows or adding rows
    n = n + 1
Wend

n = 1
numofsel = 0

While n <= p_size
    Randomize 'initialize the random number generator
    temp = Rnd
    If temp <= xvr Then
        selcross(n) = n
        numofsel = numofsel + 1
    End If
    n = n + 1
Wend  'this loop populates selcross
'++++++++++++++++++++++++++++++++++++

temp = numofsel Mod 2
If Not (temp = 0) Then
    Randomize
    choice = Int(Rnd)
    If choice = 0 Then ' delete a row
        done = False
        While Not done
            Randomize
            ranrow = Int(1 + Rnd * (p_size - 1)) ' picks a random row
            If (selcross(ranrow) <> 0) Then  'checks to see if the row exists ie is not zero
                selcross(ranrow) = 0
                done = True
            End If
        Wend
    Else ' add a row
        done = False
        While Not done
            Randomize
            ranrow = Int(1 + Rnd * (p_size - 1)) ' picks a random row
            If selcross(ranrow) = 0 Then
                selcross(ranrow) = ranrow 'this may be a problem here since it is not entirely random
                done = True
            End If
        Wend
    End If
End If
'=======
'this section cleans the zeros out of the array
n = 1
While n <= p_size
    q = 1
    Do While q <= p_size
        If Not (selcross(q) = 0) Then
            temp1(n) = selcross(q)
            selcross(q) = 0
            Exit Do
        End If
        q = q + 1
    Loop
    n = n + 1
Wend

selcross = temp1
pick_mates = selcross
'pick mates works
End Function
'
'
