Option Explicit

'declarations

'workbooks
Dim wbImport, wb2, wb3 As Workbook

'counters and utility
Public totalCount As Integer 'holds workbook max size
Public startRowPos As Integer 'for header y position
Dim hasNoValue 'holds getEOF() boolean value (to determine EOF)
'ArrayList Objects
'      0         1      2          3            4        'diff' not evaluated until after
Public customer, iccid, device_id, vol_allowed, vol_used As Object
Public customerE, iccidE, device_idE, vol_allowedE, vol_usedE, diff As Object
'will hold the evaluated data
'oversList for those who went over their data, undersList for those who used less than 75%
Public oversList, undersList As Object


'Set wb1 = ThisWorkbook
'Set wb2 = Workbooks.Open(Application.FindFile)
'Set wb3 = Workbooks.Add

'start point and main routine for our application
Sub buttonClicked()

    'MsgBox "hello world"
    'Cells(14, 7) = "hello world again"
    
    On Error Resume Next 'Error handling
    'select and import workbook
    Set wbImport = Workbooks.Open(Application.GetOpenFilename)
    'error check for wbImport error graceful exit
    If Err.number <> 0 Then
        MsgBox Err.Description & " " & "Please try again."
    End If
    
    'initialize cell position
    Cells(1, 1).Select
    
    'get user defined document size
    Dim charge: charge = True
    While charge = True
        On Error Resume Next
        totalCount = InputBox("Enter total amount of rows to parse.")
        If Err.number <> 0 Then
            MsgBox ("Please enter a numerical value. " & "[" & Err.Description & "]")
        End If
        If totalCount > 1000 Then
            MsgBox ("Please enter a number lower than 1000.")
        ElseIf totalCount = 0 Or totalCount = Null Or Not IsNumeric(totalCount) Then
            MsgBox "Invalid Input, please try again"
            Exit Sub
        Else: charge = False
        End If
    Wend
    
    'Call parseFile subroutine (at bottom of document)
    Call parseFile(0, 0, totalCount) 'parameters: y(row), x(column), totalAmount
    
    'Call printEvaluation subroutine (at bottom of document)
    Call printEvaluation
    
    'Cells(14, 7) = Columns

End Sub


'this routine will look for first cell that is not empty
'reduces need for stringent formatting
'does require (and assumes) first contact is header
Sub lookupStart(x As Integer, y As Integer)
    'declare local variables
    Dim count, countOver
    
    MsgBox "Inside lookupStart" 'debug
    'init local variables (for custom plotting)
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
    
    MsgBox "End of lookupStart"
    
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
Function getNumbers(size As Integer, iden As String)

    'variable declaration
    Dim line As String
    Dim extract As String
    Dim number As Double
    Dim ops As Boolean
    Dim nextChar As Characters
    Dim count As Integer
    Dim i As Integer
    Dim j As Variant
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
        'for loop, mid, isNumeric method (to find one number)
        For i = 1 To Len(line)
            j = Mid(line, i, 1)
            If (IsNumeric(j) = True) Then
                extract = extract & CStr(j)
            End If
            If j = "." Then
                extract = extract & CStr(j)
            End If
        Next
        number = CDbl(extract)
        
        'additional processing
        If iden = "MB" Then
            number = number * 1024 'convert GB to MB
        End If
        'add to List
        myList.Add number
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


'creates evaluated List
Function setOversList(size As Integer)

    'variable declarations
    Dim index As Integer: index = 0
    Dim tandem As Double: tandem = 0
    
    'create ArrayList Object for diff
    Set oversList = CreateObject("System.Collections.ArrayList")
    
    'evaluate vol_allowed against vol_used to calculate overages
    While index < size
    
        MsgBox "in penultimate while"
        
        On Error Resume Next
        tandem = vol_allowed.Item(index) - vol_used.Item(index)
        oversList.Add tandem
        
        If tandem > 0 Then
        
            customerE.Add customer.Item(index)
            iccidE.Add iccid.Item(index)
            device_idE.Add device_id.Item(index)
            vol_allowedE.Add vol_allowed.Item(index)
            vol_usedE.Add vol_used.Item(index)
            diff.Add tandem
            
        Next
        
        If Err.number <> 0 Then
            MsgBox Err.Description
        End If
        
        index = index + 1
        
    Wend
    
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
            Set vol_allowed = getNumbers(size, "MB")
            locked(3) = 1
        End If
        
        'vol_used data
        If LCase(ActiveCell) = "data volume (mb)" And locked(4) < 1 Then
            'set vol_used ArrayList to the return value of getNumbers function
            Set vol_used = getNumbers(size, "")
            locked(4) = 1
        End If
        
        'increment column counter
        i = i + 1
        
        'select header Row and next Column over
        Cells(startRowPos, i).Select
        ActiveCell.Offset(0, 1).Select
        
    Wend
    'create oversList ArrayList
    Set oversList = setOversList(size)
    
    'remove 'If' statement key
    Erase locked
    
End Function

'-----------------------------------------------------------------------------


'controls the process of parsing the imported file and storing data
Sub parseFile(x As Integer, y As Integer, size As Integer)
    
    Call lookupStart(x, y)
    Call lookupHeaders(size)
    
    MsgBox "Back in parseFile"
    
End Sub
'-----------------------------------------------------------------------------


'controls the process of creating the output Workbook
Sub printEvaluation()



End Sub
