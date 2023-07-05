# Excelworksheets
All the Excel worksheets where I have used the functionalities of Excel like tables, pivot tables, sorting, conditional formatting, index, match, Hlookup, vlookup, VBA and MACROS etc.

The various modules used in the files are mentioned over here.
The explanation of the various modules and functionalities.
Various formulas or shortcuts used are mentioned over here.


#Excel basic shortcuts and formulas:
1.	Zoom – Alt + V + Z   or   Ctrl + mouse scroll
2.	Total columns in Excel = 16384
3.	Ctrl + left arrow = extreme left column
4.	Ctrl + right arrow = extreme right column
5.	Total rows = 1048576
6.	Ctrl + down arrow = extreme last row
7.	Ctrl + up arrow = extreme first row
8.	New sheet = shift +f11  or   add button downside
9.	Right-click for sheet del and rename
10.	Ctrl + z = undo   and   ctrl + y = redo (only for the data entered in the sheet, not for sheet deletion and stuff)
11.	Alt + h + o + I = autofit the selected column size
12.	Alt + = -- autosum
13.	Ctrl + + and alt + I + C -- appending rows and columns
14.	Ctrl + - and Alt + I + R == deleting rows or columns
15.	For adding an ‘x’ row/column, we first select x row/column and then add them by the keywords used in 13 and 14
16.	13 & 14 both will work when the place from where the row or the column is to be deleted or added
17.	Alt + h + b + a – All borders
18.	Alt + a + t – Apply filter
19.	Custom sort for sorting with multiple conditions
20.	Absolute reference - $A$4
21.	Relative reference – A4


#Excel intermediate formulas and shortcuts:
1.	Sorting – Single-level, Multi-level, Custom sort
2.	Filter list – Auto filter tool (gets drop down in the headers)
3.	Subtotals – first sort, then convert the table into range via the table design tab and then find the subtotals.
4.	Duplicates – We first find the duplicates using the conditional formatting tab and then the duplicates can be removed by using the remove duplicates option in the table design tab.
5.	Various functions – dsum, daverage, dcount with various criteria like And, or etc.
6.	Data validation – We can use data validation to provide the users or restrict the users to enter some valid data that can be defined. We can also customize the error messages shown if wrong data is entered.
7.	Data import/export – We can import or export data easily to/from the text files, ms access files.
8.	Excel pivot tables – We can create pivot tables with clicks and drags of mouse cursors, which help us add slicers.
9.	Freeze panes – for freezing the headers of the table.
10.	Format painter – used for copying the same format of one cell or row to the other desired row or cells.
    

#Excel advance level formulas and shortcuts:
1.	Range of cells can be selected and named according to our needs for accessing them by using the names whenever and wherever needed. Ex – for calling them in functions or accessing them from anywhere in the workbook. We have to go to the formulas ribbon and then access the name manager for editing or deleting the name ranges.
2.	If - =if(condition, if true then, if false then)
3.	AND - =and(condition,condition2…), if all true then returns True else False.
4.	Nesting and within if - = if(And(conditions),statement for True, statement for false)
5.	Countif - = Countif(range of cells, Criteria)
6.	Sumif - =sumif(range where the specified condition is to be checked, the condition, range of cells where the sum is to be performed)
7.	Iferror - =iferror(value to check, if error then the statement to be shown)
8.	Vlookup - =Vlookup(lookup value, table array, column index range, range lookup)
Lookup value – the value through which the data will be retrieved from the other sheet.
Table array – the other table from which the data is to be retrieved
Column index range – the column number of the value to be found
Range lookup – either true or false: true for finding the closest match and false for finding the exact match.
9. Hlookup - =Hlookup(lookup value, table array, row index range, range lookup)
Lookup value – the value through which the data will be retrieved from the other sheet.
Table array – the other table from which the data is to be retrieved
row index range – the row number of the value to be found
Range lookup – either true or false: true for finding the closest match and false for finding the exact match.
9.	For hlookup and vlookup the limitation is that they look for the lookup value in the first column/row of the master table/sheet.
10.	Index function – returns a value at a specific position. =index(the table, rownum, colnum).
11.	Match function – returns a numeric position of a value. =match(lookup value,lookup array, match type). The match type can be either 1,0 or -1.
1 for the same value if found else a greater value. 0 for the exact same value. -1 for the same value if found else the smallest found value.
12.	The index and match functions can be used together to do the work of the hlookup and vlookup functions. The match function returns the row/col num and the index function will return the value at that row/col num.
13.	We can also use the match function with the hlookup and vlookup functions to make it dynamic.
14.	Left – = left(text, chars) 
15.	Right - = right(text, chars)
16.	Mid - = mid(text, starting index, number of chars)
17.	The left, right and mid functions can be used with the if functions for better results.
18.	Search - =search(chars to be searched, text where to be searched); returns a number that is the index of the item being searched.
19.	Concatenate/Concat - =concat(text1,text2,…) ; can concatenate 255 strings at a single go.
20.	Trace precedents – used for tracing the cells/rows used for a particular formula or in other words telling about a formula on which all cells the formula depends.
21.	Trace dependents – tells the cells which are dependent on a particular cell for its calculations.
22.	Both trace precedents and trace dependents are found in the formula tab and the arrows hence made can be removed easily by clicking on the remove arrows button.
23.	Watch window – by using the watch window option we can easily keep a watch on the selected cell added to the watch window on each and every sheet. It is also found in the formula tab.
24.	Show formula – this will help us see all the formulas used in the particular sheet. It is also present in the formula tab.
25.	Locking workbook with passwords – file > info > workbook password.
26.	WE can also protect the workbook structure and sheet data via passwords. We can find both of these options in the review panel.
27.	Goal seek in data tab – Helps us to attain a specific goal amount by changing the initial values.
28.	Solver – Add it from add-ins in the options tab. Available in the data ribbon. It is used to solve some calculations like achieving some result(max or min etc) by following the given constraints.
29.	Scenario manager in the what-if analysis section under the data ribbon lets us create multiple scenarios for a single set of data which can be changed and the changes in the results can be easily noted.


#Excel VBA and macros formula and shortcuts along with some basic codes
1.	Macros: used for performing a similar set of tasks on a table or a set of data. Like adding the headers, formatting the cells, etc. Found in the developer ribbon and starts by pressing the start recording button.
2.	VBA: used for editing the macros. If we have to make some small changes in the macro we will have to record the whole macro again but by the help of VBA we can easily edit those changes by editing the code that is present in the module.
3.	VBA stands for Visual Basic for Applications and hence can be used in most office applications.
4.	Excel consists of many objects such as the workbook, sheets, cells, tables, etc. VBA being an object-oriented programming language can be used to work on these objects.
5.	VBA helps to communicate and manipulate the macros and Excel data.
6.	If we are running a code inside the immediate window for a result then we have to start with a “?”.
7.	There are 3 types of procedures in VBA:
a.	Sub procedure: Run the code line-wise in the given sequence.
b.	Function Procedure: Run the code line-wise but returns a value too.
c.	Property procedure: We can use them to create our own defined objects and make specific procedures for these objects.
8.	For getting the resources about a function/procedure we have to keep the cursor somewhere in the function name and then press F1, it will open Microsoft’s Reading content.
9.	





VBA CODE REFERENCES:

1.	?Activeworkbook.Name
Book1
?activeworkbook.ActiveSheet.Name
Sheet1
ActiveSheet.Name = "Weekly Report"
Activesheet.range("A2").value = "Hello Aniket"
2.	To make a comment in VBA we have to use a ‘.
3.	Msgbox(prompt) – can be used to add a message box in the sheet; to get a value from a cell in  the message box – msgbox(activesheet.range(“Cell Index”).value)
4.	We can define variables inside VBA by using :
Dim varName as datatype 
Then assigning it the value by:
varName = “Value”
Code Example : 
Public Sub FunWithVariables()
    Dim userName As String
    Dim userAge As Integer
    userName = "Aniket"
    userAge = "22"
    MsgBox ("Hello " & userName & "! You are " & userAge & " years old")
    MsgBox ("In 10 years you will be " & userAge + 10 & " Years old")
    MsgBox ("You were born in " & Year(Now()) - userAge)
End Sub
5.	If else condition :
Public Sub FunWithLogic()
    If ActiveCell.Value() > 21 Then
        MsgBox ("User is older than 21")
    ElseIf ActiveCell.Value() = 21 Then
        MsgBox ("User is 21 years old")
    Else
        MsgBox ("The user is younger than 21")
    End If
End Sub

6.	Select statement :
Public Sub FunWithSelect()
    Select Case ActiveCell.Value()
        Case Is >= 21
            MsgBox ("User is older than 21")
        Case 18 To 21
            MsgBox ("the user is between 18 and 21")
        Case Else
            MsgBox ("User is less than 18 years")
    End Select
End Sub
	7.Select is better as compared to the if statement as it specifies the value of the data to be checked only once whereas in IF we have to use it multiple times.
	8. Do While loop :
		Static code:
Public Sub FunWithDoWhile()
    Dim i As Integer
    i = 1
    Do While i <= 10
        FunWithLogic
        ActiveCell.Offset(1, 0).Select
        i = i + 1
    Loop
End Sub
Dynamic Code:
Public Sub FunWithDoWhile()
    Do While ActiveCell.Value <> "" ' active cell value is not equal to an empty cell.
        FunWithLogic
        ActiveCell.Offset(1, 0).Select 'offset(1,0) defines to move by 1 row and 0 column
    Loop
End Sub


10.	For Each loop :
Public Sub FunWithForEach()
    Dim user As Range
    For Each user In Selection ‘selection defines the range that was selected in the sheet
        FunWithLogic
        ActiveCell.Offset(1, 0).Select
    Next user
End Sub
11.	For next loop:
Public Sub FunWithForNext()
    Dim i As Integer
    For i = 1 To ActiveSheet.UsedRange.Rows.Count – 1 ‘counts the number of rows and (-1) to leave the headers and start just after the headers.
        FunWithLogic
        ActiveCell.Offset(1, 0).Select
    Next i
End Sub




