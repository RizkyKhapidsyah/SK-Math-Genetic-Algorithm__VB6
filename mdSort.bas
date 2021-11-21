Attribute VB_Name = "mdSort"
Option Explicit
Option Base 1

'This module contains the code that sorts an array of chromosomes based on
'there fitness


Public Function sort(pop() As chromosome) As chromosome()
'This function will perform a simply bubble sort to sort the
'input chromosome population into ascending order (from smallest to largest)

'This could obviously be improved by using some faster means of sorting such as a
'quick sort. But to be honest with you, the GA spends most of it's time (usually)
'computing the fitness function! So by changing from an easy to understand
'sorting method to one that involves recursion probably would
'cloud the understanding of the GA
Dim i As Integer 'index
Dim j As Integer 'index
Dim upper As Integer 'the index of the last element in the pop array
Dim temp As chromosome

upper = UBound(pop)

For i = 1 To upper
    For j = i To upper
        If pop(j).fitness < pop(i).fitness Then 'swap the values
            temp = pop(j)
            pop(j) = pop(i)
            pop(i) = temp
        End If
    Next j
Next i

sort = pop
End Function
