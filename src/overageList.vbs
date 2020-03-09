Option Explicit
' Updated number logic to determine between GB and MB processing
' List Generator: for client data overages. v.1.1
' the purpose of this program is to parse an excel worksheet to find a list of
' users/clients who have gone over their data limit.
' The length (rows) of the imported list can be specified at the onset of the program,
' otherwise 300 is the default.


'declarations

'workbooks
Private wbImport, wbAdd, wbThis As Workbook
Private wbPath As String

'global locks
Private keyLock As Integer

'counters and utility
Private startRowPos As Integer 'for header 'y' position
Private globalRow As Integer 'holds current row count (variable value)
Private globalColumn As Integer 'hold current column count (fixed value)

Private declaration As String 'change in main sub
Private signature As String 'change in main sub
Private defaultAmount As Integer 'change in main sub
Private maxAmount As Integer 'change in main sub

Private borderColor As Integer 'change in the main sub
Private borderColorOwe As Integer 'change in the main sub
Private headerBackgroundColor As Integer 'change in main sub
Private headerTextColor As Integer 'change in the main sub

Private hasNoValue 'holds getEOF() boolean value (to determine EOF)
'ArrayList Objects
'       0         1      2          3            4
Private customer, iccid, device_id, vol_allowed, vol_used As Object

'oversList for those who went over their data, undersList for those who used less than 75%
'arrays to hold index locations for row retrieval
Private oversList(), undersList()



'--------------------------------------------------------------------------------------------------
'       main routine - buttonClicked() delegates tasks to parseFile() and printEvaluation()


'effectively the main routine for this application
Sub buttonClicked()
    'variable declarations
    Dim totalCount As Variant
    Dim totalCountInt As Integer: totalCountInt = 0
    Dim charge As Boolean
    Dim lockBox As Integer
    
    'set application workbook as ThisWorkBook
    wbPath = ThisWorkbook.Path
    
    '[<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'global values can be changed here
    
    'logic
    globalColumn = 7 ' change column count for whole document
    
    'formatting
    declaration = "Clients who have gone over their data limit are listed below" 'bold large text at top of document
    signature = "For more information contact the I.T. Department" 'smaller message below declaration
    defaultAmount = 300 'default amount of rows to parse (more rows takes longer - difficulty is high due to recursive parse)
    maxAmount = 1001 'maximum amount (to negate chance of mistaken input and resulting undesirable wait time)
    
    borderColor = 35 'change border color here
    borderColorOwe = 3 'change arrears border color here
    headerBackgroundColor = 0 'change header background color
    headerTextColor = 1 'change header text color here, e.g. change to white (2) if borderColor is black (1)
    
    ']<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    
    'MsgBox "hello world"
    Range("B10", "D10").Interior.ColorIndex = 35
    Cells(10, 2) = "Program is running..."
    
    On Error Resume Next 'Error handling
    'select and import workbook
    
    While lockBox < 1
        Set wbImport = Workbooks.Open(Application.GetOpenFilename)
        
        'if no file is chosen then exit
        If Err.number = 1004 Then
            MsgBox "You didn't choose anything, Goodbye!" & " [ " & Err.Description & " ] "
            Call cleanup
            Exit Sub
        
        'error check for wbImport error graceful exit
        ElseIf Err.number <> 0 Then
            MsgBox Err.Description & " " & "Please try again."
            
        ElseIf Err.number = 0 Then
            lockBox = lockBox + 1
        End If
    Wend
    
    'initialize cell position
    Cells(1, 1).Select
    
    'get user defined document size
    
    charge = True
    While charge = True
        'variable declarations
        lockBox = 0
        
        'get user input
        totalCount = InputBox("Enter number of rows to search or PRESS ENTER for DEFAULT (" & defaultAmount & "). To QUIT type 'X' or 'Q'.")
        
        If totalCount = "" Or totalCount = "\n" Then
            MsgBox "Default amount set:" & " [ " & CStr(defaultAmount) & " " & "Max Rows" & " ] "
            totalCount = defaultAmount
            
        ElseIf LCase(totalCount) = "x" Or LCase(totalCount = "q") Or totalCount = 0 Then
            Call cleanup 'quit and refresh
            Exit Sub
        
        ElseIf (Not IsNumeric(totalCount)) And totalCount <> "\n" Then
            MsgBox "Please enter a numerical value. " & " [ " & Err.Description & " ] "
            lockBox = lockBox + 1 'lock the lockBox
        
        ElseIf totalCount > maxAmount Then 'for maxAmount
            MsgBox ("Please enter a number lower than " & maxAmount & ".")
            lockBox = lockBox + 1 'lock the lockBox
            
        ElseIf IsEmpty(totalCount) Then
            Call cleanup
            Exit Sub
        End If
        
        'unlock if true
        If lockBox = 0 Then
            charge = False
        End If
        'if nothing goes wrong, lockbox is open and loop exits
    Wend
    
    'store variant inside integer to avoid overflow
    totalCountInt = CInt(totalCount)
    
     'debug block code
    'On Error Resume Next
    
    'Call parseFile subroutine (at bottom of document)
    Call parseFile(0, 0, totalCountInt) 'parameters: y(row), x(column), totalAmount
       
    'If Err.number <> 0 Then
    '    MsgBox "#1" & Err.Description
    '    MsgBox Err.number
    '    Err.Clear
    'End If
    
    'Call printEvaluation subroutine (at bottom of document)
    Call printEvaluation(totalCountInt)
    
    'If Err.number <> 0 Then
    '    MsgBox "#2" & Err.Description
    '    MsgBox Err.number
    '    Err.Clear
    'End If
    
    'washing dishes time! *quietly leaves
    Call cleanup
    
<<<<<<< HEAD
=======
    'If Err.number <> 0 Then
    '    MsgBox "#1" & Err.Description
    '    MsgBox Err.number
    '    Err.Clear
    'End If
    
>>>>>>> cadanu
    

End Sub 'end of main routine
    

'release script resources
Function cleanup()
    
    'uninitialize non-object variables
    keyLock = Empty
    startRowPos = Empty
    globalRow = Empty
    globalColumn = Empty
    borderColor = Empty
    borderColorOwe = Empty
    headerTextColor = Empty
    headerBackgroundColor = Empty
    
    'remove object pointers
    Set hasNoValue = Nothing
    Set customer = Nothing
    Set iccid = Nothing
    Set device_id = Nothing
    Set vol_allowed = Nothing
    Set vol_used = Nothing
    Erase oversList()
    Erase undersList()
    'close imported Workbook after use
    wbImport.Close
    
    'close application Workbook after use
    ThisWorkbook.Activate
    Cells(10, 2) = "Program will exit..."
    MsgBox "Program run complete." & " [ " & "Exit code: " & Err.number & " ] "
    Cells(10, 2) = "Program will exit..."
    ThisWorkbook.Close (False)
        
    'debug block code
    'On Error Resume Next
    'If Err.number <> 0 Then
    '    MsgBox Err.Description
    '    MsgBox Err.number
    '    Err.Clear
    'End If
        
End Function



'                               parseFile() processes below this line
'--------------------------------------------------------------------------------------------------



'this routine will look for first cell that is not empty
'reduces need for stringent formatting
'does require (and assumes) first contact is header
Sub lookupStart(x As Integer, y As Integer)

    'declare local variables
    Dim count, countOver
    'init local variables (for custom plotting set x,y in parseFile())
    count = y
    countOver = x
    
    'select top left corner
    Range("a1").Select
    
    'if the cell is empty then keep looking (right 9, down 9)
    While ActiveCell = Null Or ActiveCell = "" And count < 9
    
        'first, go over (right) and look
        While ActiveCell = Null Or ActiveCell = "" And countOver < 9
            ActiveCell.Offset(0, 1).Select
            countOver = countOver + 1
        Wend
        
        'then go down one if above unsuccessful
        If ActiveCell = Null Or ActiveCell = "" Then
            ActiveCell.Offset(1, -countOver).Select
        End If
        'increment, repeat
        count = count + 1
        'empty countOver
        countOver = x
        'if successful this while will exit
    Wend
    
    'prevents count from being 0
    If count < 1 Then
        count = 1
    End If
    'sets public startRowPos variable as count variable value
    startRowPos = count
    
End Sub


'this function looks for the next header
Sub lookupNext(indexBool As Boolean)

    'select next
    If indexBool = False Then
        MsgBox "I'm in ToDo"
        ActiveCell.Offset(0, 1).Select
    End If
    
End Sub


'function for EOF parsing (looks ahead 10 cells)
Function getEOF(ByVal place As Range)
    'local variables
    Dim eofCounter: eofCounter = 0
    Dim eofValue: eofValue = 0
    'evaluates if ActiveCell (place) is empty and the 10 following cells are also empty
    While place = Null Or place = "" And eofCounter < 9
        place.Offset(1, 0).Select
        If place = Null Or place = "" Then
            eofValue = eofValue + 1
        End If
        eofCounter = eofCounter + 1
    Wend
    If eofCounter > 8 Then 'if more than 9 empty cells recorded
        hasNoValue = True
    Else: hasNoValue = False
    End If
    getEOF = hasNoValue 'return boolean
End Function


'evaluates "monthly rate plan", and "" columns and extracts numerical values
Function getNumbers(size As Integer) ', iden As String

    'variable declaration
    Dim line As String
    Dim extract As String
    Dim converse As String
    Dim number As Double
    Dim letters As String
    Dim ops As Boolean
    Dim nextChar As Characters
    Dim count As Integer
    Dim i As Integer
    Dim j As Variant
    Dim k As Variant
    Dim m As Integer 'loop index modifier
    'initialize ArrayList location in memory
    Dim myList
    
    count = 0
    'create new ArrayList in memory location
    Set myList = CreateObject("System.Collections.ArrayList")
    ActiveCell.Offset(1, 0).Select 'step off header
    
    While count < size 'fill ArrayList 'list'
        'init line then assign value of ActiveCell
        line = ""
        line = ActiveCell.Value
        
        'if loop to prevent 'type' error if cell is empty or null
        If Len(line) > 0 Or line = Null Then
            'for loop, mid, isNumeric method (to find one number)
            
            For i = 1 To Len(line)
            
                If i > 1 Then
                    m = i - 1
                    k = Mid(line, m, 1) 'for previous char matching
                End If
                
                j = Mid(line, i, 1) 'j = mid function (string, start, length)
                
                'test for numbers
                If (IsNumeric(j) = True) Then
                    extract = extract & CStr(j)
                End If
                
                'test for decimals
                If j = "." Then
                    extract = extract & CStr(j)
                End If
                
                'test for conversion rate (GB or MB)
                If LCase(k) = "g" Or LCase(k) = "m" And LCase(j) = "b" Then
                    converse = CStr(k) & CStr(j)
                End If
                
            Next
            number = CDbl(extract)
            letters = CStr(converse)
            
        Else
            number = CDbl(0)
            letters = ""
        End If
        
        'process for GB or MB
        'If LCase(letters) = "gb" Then
        '    iden = "gb"
        'End If
        
        'additional processing
        If LCase(letters) = "gb" Then
            number = number * 1024 'convert GB to MB
        End If
        
        'add to List
        myList.Add Round(number, 1)
        'update Cell position
        ActiveCell.Offset(1, 0).Select
        'update counter and empty variables
        count = count + 1
        j = Null
        extract = ""
        
    Wend
    Set getNumbers = myList 'return list
    
End Function


'fills ArrayList with column data
Function getHeaderData(size As Integer) As Object

    'variable declarations
    Dim count: count = 0
    'create memory location for ArrayList reference variable
    Dim myList
    'create ArrayList Object
    Set myList = CreateObject("System.Collections.ArrayList")
    ActiveCell.Offset(1, 0).Select 'step off header
    
    While count < size 'fill ArrayList 'list'
        myList.Add ActiveCell
        ActiveCell.Offset(1, 0).Select
        count = count + 1
    Wend
    Set getHeaderData = myList 'return list
    
End Function


'this routine will lookup headers to get data from column
'perhaps this is the most important function
Function lookupHeaders(size As Integer)

    'determines whether ActiveCell is origin
    Dim locked(4) 'locks if statement after entry
    Dim i As Integer: i = 0
    Dim line As Variant
    
    'this while loop ends on an empty cell
    While ActiveCell <> Null Or ActiveCell <> ""
    
        'customer data
        If LCase(ActiveCell) = "customer" And locked(0) < 1 Then
            Set customer = getHeaderData(size)
            'MsgBox "got 1"
            locked(0) = 1
        End If
        
        'iccid data
        If LCase(ActiveCell) = "iccid" And locked(1) < 1 Then
            Set iccid = getHeaderData(size)
            'MsgBox "got 2"
            locked(1) = 1
        End If
        
        'device_id data
        If LCase(ActiveCell) = "device id" And locked(2) < 1 Then
            Set device_id = getHeaderData(size)
            'MsgBox "got 3"
            locked(2) = 1
        End If
        
        'vol_allowed data
        If LCase(ActiveCell) = "monthly rate plan" And locked(3) < 1 Then
            'set vol_allowed ArrayList to the return value of getNumbers function
            Set vol_allowed = getNumbers(size)
            locked(3) = 1
        End If
        
        'vol_used data
        If LCase(ActiveCell) = "data volume (mb)" And locked(4) < 1 Then
            'set vol_used ArrayList to the return value of getNumbers function
            Set vol_used = getNumbers(size)
            locked(4) = 1
        End If
        
        'increment column counter
        i = i + 1
        
        'select header Row and next Column over
        Cells(startRowPos, i).Select
        ActiveCell.Offset(0, 1).Select
        
    Wend
    'create oversList ArrayList
    'Set oversList = setOversList(size)
    
    'remove 'If' statement key
    Erase locked
    
End Function



'--------------------------------------------------------------------------------------------------
'                               parseFile() processes above this line


'controls the process of parsing the imported file and storing data
Sub parseFile(x As Integer, y As Integer, size As Integer)
    
    Call lookupStart(x, y)
    Call lookupHeaders(size)
    
End Sub


'                               printEvaluation() processes below this line
'--------------------------------------------------------------------------------------------------



'creates evaluated List
Function setOversList(size As Integer)

    'variable declarations
    Dim counter As Integer
    Dim index As Integer: index = 0
    Dim tandem As Double: tandem = 0
    
    'evaluate vol_allowed against vol_used to get array size
    While index < size
        tandem = CDbl(vol_used.Item(index)) - CDbl(vol_allowed.Item(index))
        If tandem > 0 Then
            counter = counter + 1
        End If
        index = index + 1
    Wend
    
    'set array size
    ReDim oversList(counter)
    
    index = 0
    counter = 0
    'evaluate vol_allowed against vol_used to calculate overages
    While index < size
        tandem = CDbl(vol_used.Item(index)) - CDbl(vol_allowed.Item(index))
        If tandem > 0 Then
            oversList(counter) = index
            counter = counter + 1
        End If
        index = index + 1
    Wend
    'at this point, array will contain index positions of rows that match requirements
    
End Function


Function printer(size As Integer)
    
    'variable declaration
    Dim index As Integer
    Dim indexOvers As Integer
    'declare and define variables for calculations
    Dim amountOver As Double
    Dim percentOver As Double
    Dim percent As String
    
    keyLock = 0
    'this while loop prints headers
    While keyLock < 1
        
        ActiveCell.Interior.ColorIndex = headerBackgroundColor
        ActiveCell.Font.size = 14
        ActiveCell.Font.ColorIndex = headerTextColor
        ActiveCell.Font.Bold = True
        ActiveCell.EntireColumn.AutoFit
        ActiveCell = "Customer"
        ActiveCell.Offset(0, 1).Select
        
        ActiveCell.Interior.ColorIndex = headerBackgroundColor
        ActiveCell.Font.size = 14
        ActiveCell.Font.ColorIndex = headerTextColor
        ActiveCell.Font.Bold = True
        ActiveCell.EntireColumn.AutoFit
        ActiveCell = "ICCID"
        ActiveCell.Offset(0, 1).Select
        
        ActiveCell.Interior.ColorIndex = headerBackgroundColor
        ActiveCell.Font.size = 14
        ActiveCell.Font.ColorIndex = headerTextColor
        ActiveCell.Font.Bold = True
        ActiveCell.EntireColumn.AutoFit
        ActiveCell = "Device ID"
        ActiveCell.Offset(0, 1).Select
        
        ActiveCell.Interior.ColorIndex = headerBackgroundColor
        ActiveCell.Font.size = 14
        ActiveCell.Font.ColorIndex = headerTextColor
        ActiveCell.Font.Bold = True
        ActiveCell.EntireColumn.AutoFit
        ActiveCell = "Monthly Rate Plan"
        ActiveCell.Offset(0, 1).Select
        
        ActiveCell.Interior.ColorIndex = headerBackgroundColor
        ActiveCell.Font.size = 14
        ActiveCell.Font.ColorIndex = headerTextColor
        ActiveCell.Font.Bold = True
        ActiveCell.EntireColumn.AutoFit
        ActiveCell = "Data Used (MB)"
        ActiveCell.Offset(0, 1).Select
        
        ActiveCell.Interior.ColorIndex = headerBackgroundColor
        ActiveCell.Font.size = 14
        ActiveCell.Font.ColorIndex = headerTextColor
        ActiveCell.Font.Bold = True
        ActiveCell.EntireColumn.AutoFit
        ActiveCell = "Amount Over (MB)"
        ActiveCell.Offset(0, 1).Select
        
        ActiveCell.Interior.ColorIndex = headerBackgroundColor
        ActiveCell.Font.size = 14
        ActiveCell.Font.ColorIndex = headerTextColor
        ActiveCell.Font.Bold = True
        ActiveCell.EntireColumn.AutoFit
        ActiveCell = "~ Percent Over"
        
        'lock while loop from entry
        keyLock = keyLock + 1
        're-initialize ActiveCell
        ActiveCell.Offset(1, -6).Select
        
    Wend
    
    'MsgBox "After first while"
    
    indexOvers = 0
    'this while loop prints values under headers (index value = oversList(indexOvers) value)
    While index < size 'Or index > UBound(oversList)
        
        'if loop to prevent error if cell is empty or null
        If vol_allowed(index) <> "" Or vol_used(index) <> "" Then
            amountOver = CDbl(Round(vol_used.Item(index) - vol_allowed(index), 1))
            ' to avoid x/0 overflow
            If amountOver <> 0 Then
                percentOver = CDbl(Round((amountOver * 100) / vol_allowed(index), 0))
            End If
            percent = CStr(percentOver & "%") '4
        Else
            amountOver = 0: percentOver = 0: percent = "" '5
        End If
        
        'if 'index' value is equal to value stored at 'oversList(indexOvers)'
        If index = oversList(indexOvers) Then
        
            'print out list values according to index locations
            ActiveCell.Interior.ColorIndex = 34
            ActiveCell.EntireColumn.AutoFit
            ActiveCell = customer.Item(index) '
            ActiveCell.Offset(0, 1).Select
            
            ActiveCell.Interior.ColorIndex = 19
            ActiveCell.EntireColumn.AutoFit
            ActiveCell = iccid.Item(index) '
            ActiveCell.Offset(0, 1).Select
            
            ActiveCell.Interior.ColorIndex = 19
            ActiveCell.EntireColumn.AutoFit
            ActiveCell = device_id.Item(index) '
            ActiveCell.Offset(0, 1).Select
            
            ActiveCell.Interior.ColorIndex = 19
            ActiveCell.EntireColumn.AutoFit
            ActiveCell = vol_allowed.Item(index) '
            ActiveCell.Offset(0, 1).Select
            
            ActiveCell.Interior.ColorIndex = 19
            ActiveCell.EntireColumn.AutoFit
            ActiveCell = vol_used.Item(index) '
            ActiveCell.Offset(0, 1).Select
            
            ActiveCell.Interior.ColorIndex = 22
            ActiveCell.Font.Bold = True
            ActiveCell.EntireColumn.AutoFit
            ActiveCell = amountOver
            ActiveCell.Offset(0, 1).Select
            
            ActiveCell.Interior.ColorIndex = 22
            ActiveCell.EntireColumn.AutoFit
            ActiveCell = percent
            
            're-initialize ActiveCell
            ActiveCell.Offset(1, -6).Select
            'increment overs counter
            indexOvers = indexOvers + 1
            
        End If
        'increment list counter
        index = index + 1
        
    Wend
    
    'border start location
    ActiveCell.Offset(1, -1).Select
    
    'print border around table
    Call borderPrinter
    
End Function


'prints a border around table
Function borderPrinter()

    'variable declarations
    Dim index As Integer
    Dim arrLock(3)
    
    index = 0
    'print right, up, left and down (left, right fixed : up, down variabled)
    While index < 4
    
        While arrLock(index) < globalColumn + 2 'columns right
            If arrLock(index) < 6 Then
                ActiveCell.Interior.ColorIndex = borderColor
            Else:
                ActiveCell.Interior.ColorIndex = borderColorOwe
            End If
            ActiveCell.Offset(0, 1).Select
            arrLock(index) = arrLock(index) + 1
        Wend
        index = index + 1
    
        While arrLock(index) <= UBound(oversList) + 3 'rows up
            ActiveCell.Interior.ColorIndex = borderColorOwe
            ActiveCell.Offset(-1, 0).Select
            arrLock(index) = arrLock(index) + 1
        Wend
        index = index + 1
    
        While arrLock(index) < globalColumn + 3 'columns left
            If arrLock(index) < 4 Then
                ActiveCell.Interior.ColorIndex = borderColorOwe
            Else
                ActiveCell.Interior.ColorIndex = borderColor
            End If
            ActiveCell.Offset(0, -1).Select
            arrLock(index) = arrLock(index) + 1
        Wend
        index = index + 1
    
        While arrLock(index) <= UBound(oversList) + 4 'rows down
            ActiveCell.Interior.ColorIndex = borderColor
            ActiveCell.Offset(1, 0).Select
            arrLock(index) = arrLock(index) + 1
        Wend
        index = index + 1
    
    Wend
    
End Function




'-----------------------------------------------------------------------------
'                               printEvaluation() processes above this line



'controls the process of creating the output Workbook
Sub printEvaluation(size As Integer)

    'create index of rows with overs
    Call setOversList(size)
    
    'add a new workbook for data output
    Set wbAdd = Workbooks.Add
    
    'initialize main header location and add formatting
    Cells(2, 2).Select
    ActiveCell.Value = declaration
    ActiveCell.Font.size = 24
    ActiveCell.Font.Bold = True
    
    'modify signature
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = signature '<<< Modify signature here <<<
    
    'initialize table location (top left)
    ActiveCell.Offset(4, 1).Select
    
    'printer function prints data to new Workbook
    Call printer(size)

End Sub



'-----------------------------------------------------------------------------
'                               end of document



'Author: Gordon Joyce
