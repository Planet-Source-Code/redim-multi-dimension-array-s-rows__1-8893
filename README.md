<div align="center">

## Redim multi\-dimension array's rows


</div>

### Description

At times I need to redim a multi-dimension array by adding an extra row or rows, and keep the arrays contents. In VB you can only change the size of the last dimension in a multi-dimension array, this is ok if you want to only add columns. If you want to add a row to a multi-dimension array you can’t redim it and preserve the array’s data. To get around this problem I have created a process to do this for me. Just copy the contents of the multi-dimension array to another exact copy, multi-dimension temp array. Then redim the original array’s row up by one, and then copy the temp array back to the original. Now the original multi-dimension array still has its contents plus has an extra new row to add more data to!
 
### More Info
 
Input array.

Need to know the facts about building and using arrays.

A redim multi-dimension array.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/redim-multi-dimension-array-s-rows__1-8893/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
  Dim WorkAry() As String
  Dim row As Integer, col As Integer, rowsize As Integer
  rowsize = 5
  ReDim WorkAry(rowsize, 5)
  For row = 0 To 5
   For col = 0 To 5
     WorkAry(row, col) = row & "-" & col
   Next col
  Next row
  rowsize = rowsize + 1
  Call Redim_Array(WorkAry(), rowsize)
'** now add data into the extra line for WorkAry() array. **
  col = 0
  For col = 0 To 5
   WorkAry(rowsize, col) = rowsize & "-" & col
  Next col
End Sub
Private Sub Redim_Array(WrkAry() As String, NewRowSize As Integer)
'** Redim a multi-dimension array that will allow an extra row to be added.
  Dim TempAry() As String
  Dim row As Integer, col As Integer, CurRows As Integer
'** Arrays look like this, Ary(Row, Col) with rows first then columns. **
  CurRows = NewRowSize - 1  '** need to get WrkAry() current row number. **
  ReDim TempAry(CurRows, 5) '** create same size temp array as in coming WrkAry() array. **
               '** the columns will stay the same. **
 '** move multi-dimension WrkAry() to an exact copy multi-dimension TempAry(). **
  For row = 0 To CurRows
   For col = 0 To 5
     TempAry(row, col) = WrkAry(row, col)
   Next col
  Next row
  ReDim WrkAry(NewRowSize, 5) '** re-dimension WrkAry() with one more row. **
'** copy TempAry() to WrkAry() which is now one row larger but not being used at this time. **
  For row = 0 To CurRows
   For col = 0 To 5
     WrkAry(row, col) = TempAry(row, col)
   Next col
  Next row
'** WrkAry() will keep all of its original data and has one more row for more data later. **
End Sub
```

