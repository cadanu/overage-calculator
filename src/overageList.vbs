Option Explicit

'declarations
'Dim index As Integer
'worksheets
Dim overs As Worksheet
'workbooks
Dim wbImport, wb2, wb3 As Workbook
'counters and utility
Public count, countOver, totalCount 'holds workbook max size
Dim hasNoValue 'holds getEOF() boolean value (to determine EOF)
Dim join
Public found(4) 'stores 5 columns (ArrayList) by order below
'ArrayList Objects
'       0       1       2           3           4      'diff' not evaluated until end
Dim customer, iccid, device_id, vol_allowed, vol_used, diff As Object
'Range object
Dim startPos As Range


'Set wb1 = ThisWorkbook
'Set wb2 = Workbooks.Open(Application.FindFile)
'Set wb3 = Workbooks.Add
Sub buttonClicked()

    'MsgBox "hello world"
    Cells(14, 7) = "hello world again"
    Set wbImport = Workbooks.Open(Application.GetOpenFilename)
    'wbImport.Activate
    
    Cells(1, 1).Select
    Call parseFile(0, 0, 100) 'set start point (cell) here <<
    
    'Cells(14, 7) = Columns
    

End Sub


'this routine will look for first cell that is not empty
'reduces need for stringent formatting
'does require (and assumes) first contact is header
Sub lookupStart(x As Integer, y As Integer)

    MsgBox "Inside lookupStart"
    
    count = y
    countOver = x
    
    'select top left corner
    Range("a1").Select
    
    'if the cell is empty then keep looking (right 9, down 9)
    While ActiveCell = Null Or ActiveCell = "" And count < 9
    
        MsgBox "inside lookupStart first While"
    
        'first, go over (right) and look, then go back to start
        While ActiveCell = Null Or ActiveCell = "" And countOver < 9
        
            MsgBox "inside lookupStart second While"
            
            ActiveCell.Offset(0, 1).Select
            countOver = countOver + 1
        Wend
        
        'go down one and look if above unsuccessful
        If ActiveCell = Null Or ActiveCell = "" Then
            ActiveCell.Offset(1, 0).Select
        End If
        'increment, repeat
        count = count + 1
        'if successful this while will exit
        Cells(1, count).Select
        
    Wend
    
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


'fills ArrayList with column data
Function getHeaderData(size As Integer) As Object
    'debug
    MsgBox "In getHeaderData"
    
    count = 0
    'create new ArrayList
    Dim myList
    Set myList = CreateObject("System.Collections.ArrayList")
    ActiveCell.Offset(1, 0).Select 'step off header
    While count < size 'fill ArrayList 'list'
        
        MsgBox "Inside getHeaderData While"
    
        myList.Add ActiveCell
        ActiveCell.Offset(1, 0).Select
        count = count + 1
    Wend
    Set getHeaderData = myList 'return list
End Function


'this routine will lookup headers to get data from column
Function lookupHeaders(size As Integer)
    'determines whether ActiveCell is origin
    Dim done As Boolean
    Dim i As Integer: i = 0
    Dim failSafe As Integer
    done = True
    
    While ActiveCell <> Null Or ActiveCell <> ""
        
        'customer data
        If LCase(ActiveCell) = "customer" And found(0) < 1 Then
            Set customer = getHeaderData(size)
            MsgBox "got 1"
            found(0) = 1
            i = i + 1
        End If
        
        'iccid data
        If LCase(ActiveCell) = "iccid" And found(1) < 1 Then
            Set iccid = getHeaderData(size)
            MsgBox "got 2"
            found(1) = 1
            i = i + 1
        End If
        
        'device_id data
        If LCase(ActiveCell) = "device id" And found(2) < 1 Then
            Set device_id = getHeaderData(size)
            MsgBox "got 3"
            found(2) = 1
            i = i + 1
        End If
        
        'vol_allowed data
        If LCase(ActiveCell) = "monthly rate plan" And found(3) < 1 Then
            Set vol_allowed = getHeaderData(size)
            MsgBox "got 4"
            found(3) = 1
            i = i + 1
        End If
        
        'vol_used data
        If LCase(ActiveCell) = "data volume (mb)" And found(4) < 1 Then
            Set vol_used = getHeaderData(size)
            MsgBox "got 5"
            found(4) = 1
            i = i + 1
        End If
        
        Cells(count, i + 1).Select
        
        'evaluate what has been found, add found() indexes
        'i = 0
        'While i < 5
        '    If found(i) < 1 Then
        '    MsgBox "Inside eval while-if: " & done
        '        done = False
        '    End If
        '    i = i + 1
        '    MsgBox "Inside eval while: " & i
        'Wend
        
        'if any indexes are 0, find headers (must know header amount - fixed)
        'If done = False And ActiveCell <> Null And ActiveCell <> "" Then
        '    MsgBox "Inside call lookupStart"
        '    Call lookupNext(False)
        'End If
        
        'debug
        'MsgBox "inside lookupHeaders while loop"
        'failSafe = failSafe + 1
        
    Wend
    
End Function


'controls the process of parsing the imported file for data and evaluating
Sub parseFile(x As Integer, y As Integer, size As Integer)

    Call lookupStart(x, y)
    Call lookupHeaders(size)
    
End Sub
