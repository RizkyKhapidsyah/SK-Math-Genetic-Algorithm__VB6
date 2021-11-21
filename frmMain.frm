VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Valued Genetic Algorithm"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAttributes 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Text            =   "0.05"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtAttributes 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Text            =   "0.8"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtAttributes 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Text            =   "50"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtAttributes 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Text            =   "100"
      Top             =   360
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtOutput 
      Height          =   3255
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Optimize"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Mutation Probability"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Crossover Probability"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Generations"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Population size"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1 'this is used to ensure that all arrays start at 1!
            'in any other language I don't really care, but in VB
            'the starting element index can be a bit mucked up.....
            'so it is better to explicitly state the starting index


'=============================
'Note: The optimization goal of this GA is MINIMIZATION
'       This means that if we want to maximize we simple scale the raw fitness by multiplying it by -1

Private Sub main()

Dim popSize As Integer 'number of chromosomes in a population
Dim numVar As Integer 'number of variables that each chromosome contains which should be the number that the fitness function requires
Dim numGen As Long 'number of generations to evolve the population
Dim rng() As varRange 'holds the range of the variables
Dim targetValue As Single  'the target value of the GA
Dim mutationProb As Single 'the mutation probability
Dim crossProb As Single 'the crossover probability
'==============================

Dim myPop() As chromosome 'the chromosome population
Dim i As Integer 'index for odd looping jobs
Dim gaCount As Integer 'index to keep track of the evolution
Dim selection() As Integer 'This is used to hold the indexes of the chromosomes that are selected for mating
Dim best As chromosome 'this holds the best individual that has been found so far
Dim genFound As Long 'this is used to hold the generation that the best was found

                       
'===============================
popSize = txtAttributes(0).Text
numVar = 3
numGen = txtAttributes(1).Text
targetValue = 28
mutationProb = txtAttributes(3).Text
crossProb = txtAttributes(2).Text

ReDim rng(1 To numVar) 'make sure there is room to hold the information for all the variables
rng(1).lower = -100
rng(1).upper = 100
rng(2).lower = -100
rng(2).upper = 100
rng(3).lower = -100
rng(3).upper = 100


pb.Max = numGen 'set the progress bar so that it updates properly
'==============================
myPop = popBuilder(popSize, numVar, rng) 'create the initial population
'=========================================

best = myPop(1) 'initialize the best chromosome so that it won't produce an error
best.fitness = 10000000 'Initialize the fitness to a large number that you know will not come up
                        'The reason for doing this is too make sure that the smallest value
                        'of the first generation gets selected. Also, if you do not manually
                        'initialize a variable, it will default to zero, our goal!!!
                        
genFound = 0 'This is not really necessary as VB defaults to zero. I do this because
             'it is much clearer when you read the source code what it is supposed to start at

'The main loop start here, ie the loop that counts the number
'of generations to evolve the population for

For gaCount = 1 To numGen
'=================================

'evaluate the population
For i = 1 To popSize
    myPop(i).value = fitness(myPop(i).dv) 'remember assign the raw fitness to the value variable first
    myPop(i).fitness = scaleValue(myPop(i).value, targetValue)
Next i
'The reason that I choose to loop through the population and evaluate
'the fitness in the above manner is simple, if you have a
'complicated fitness function you don't have to worry about
'iterating the population through it, you just have to
'worry about coding the function itself....
'You could always put the snippet of code into it's own module to improve readablity
'=========================================

'sort the population so that the smallest value is in the first index position
'sorting the population makes it relatively simply to find the fittest
'chromosomes, ie they are at the top....
'alternatively, depending on the information that you want to display,
'you could find only the smallest one of the population which would
'significantly speed this up. I didn't do this because I wanted to
'produce routines that were as general as possible.....

'Remember that the sorting routine could be changed to a quick sort algorithm.
'The sorting algorithm that I used is a bubble sort type algorithm....I used this
'because I can not assume that anyone who reads this would be able to understand
'what is going on as the quicksort routine employes recursion and although I understand
'what recursion is, I don't have a solid grasp on the quick sort algorithm....

'It is mostly a matter of preference... In my experience and the problems that
'I dealt with, the sorting was not the slowest link in the chain.... but for simple
'problems like this, it is the slowest link in the chain....

myPop = sort(myPop)

'====================
'The below code is used to display the values of all the chromosomes from generation
'to generation
'record the best individual....ie the one at the begining of the list....
'For i = 1 To popSize
'    printChrom myPop(i)
'Next i

'figure out if the best chromosome for the current generation is the best one
'found so far
If myPop(1).fitness < best.fitness Then
    best = myPop(1)
    genFound = gaCount
End If
'=====================
'Some kind of termination procedure can be implemented here.....
'In my experience, a termination procedure does not work very well because in
'most cases the optimal solution is unknown so it is difficult to set a termination
'criteria.... but in this case you could use: if best.fitness < 0.00001 then exit sub....

'======================

'display progress.....
DoEvents 'this allows the text box to be updated
If gaCount Mod 10 = 0 Then 'only display the stats every 10th generation

txtOutput.Text = txtOutput.Text & "Current Generation: " & gaCount & vbNewLine
txtOutput.Text = txtOutput.Text & "The Best Chromosome found so far is:" & vbNewLine
txtOutput.Text = txtOutput.Text & "dv 1:" & best.dv(1) & vbNewLine
txtOutput.Text = txtOutput.Text & "dv 2:" & best.dv(2) & vbNewLine
txtOutput.Text = txtOutput.Text & "dv 3:" & best.dv(3) & vbNewLine
txtOutput.Text = txtOutput.Text & "Value:" & best.value & vbNewLine
txtOutput.Text = txtOutput.Text & "Fitness:" & best.fitness & vbNewLine
txtOutput.Text = txtOutput.Text & "goal:" & targetValue & vbNewLine
txtOutput.Text = txtOutput.Text & "generation:" & genFound & vbNewLine
txtOutput.Text = txtOutput.Text & "=======================" & vbNewLine
End If

'=========================================
'selection- This section selects individuals from the population
'           that will go onto the crossover population

myPop = tournament(myPop, 4)
'=========================================
'crossover the selected population

'first select the chromosomes that will undergo crossover
selection = pick_mates(crossProb, popSize)

myPop = xover(myPop, selection, rng)

'=========================================
'mutate the entire population

myPop = mutation(myPop, mutationProb, rng)

'========================================
pb.value = pb.value + 1
Next gaCount

End Sub

Private Sub cmdCalculate_Click()
    pb.value = pb.Min 'make sure the progress bar's value is zero
    txtOutput.Text = "" 'clear the textbox
    main
End Sub

Private Sub Form_Load()
    pb.value = 0
    pb.Min = 0
    pb.Max = 100
End Sub

Private Sub printChrom(indiv As chromosome)

'this sub routine is used to print out the chromosome values
'I used this to debug some of the routines.....

txtOutput.Text = txtOutput.Text & "dv 1:" & indiv.dv(1) & vbNewLine
txtOutput.Text = txtOutput.Text & "dv 2:" & indiv.dv(2) & vbNewLine
txtOutput.Text = txtOutput.Text & "dv 3:" & indiv.dv(3) & vbNewLine
txtOutput.Text = txtOutput.Text & "Value:" & indiv.value & vbNewLine
txtOutput.Text = txtOutput.Text & "Fitness:" & indiv.fitness & vbNewLine
txtOutput.Text = txtOutput.Text & "=======================" & vbNewLine


End Sub

Private Sub pb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
