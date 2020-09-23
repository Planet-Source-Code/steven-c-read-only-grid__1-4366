VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**** DATABASE VARIABLES
Private mvarDataBase As Database
Private mvarRecordset As Recordset

'**** ARRAY VARIABLES
Private mvarDisplay() As Variant          'Array that controls actual display
Private mvarFldNames() As String
Private mvarTextBoxes() As TextBox

'**** FORM CONTROL VARIABLES
Private mvarScrollBar As FlatScrollBar
Private mvarProgressMeter As ProgressBar


'**** OTHER VARIABLES
Private mvarMaxRecords As Long
Private mvarRecPosition As Long

Private mvarCols As Integer             'Total Number of Columns
Private mvarCol As Integer              'Current Column Position
Private mvarRows As Integer             'Total Number of Rows
Private mvarRow As Integer              'Current Row Position

Public Property Let DatabaseName(ByVal vData As String)
    
    Dim sMsg As String
    Dim sTitle As String
    
    '*** Database name that contains info to be displayed.
    
    On Error GoTo DbOpenError
    Set mvarDataBase = OpenDatabase(vData)
    
    Exit Property
    
DbOpenError:
    sTitle = "Grid Class Database Open Error"
    sMsg = Err.Number & " Occurred." & vbCr & Err.Description
    MsgBox Prompt:=sMsg, Buttons:=vbOKOnly + vbCritical, Title:=sTitle
    Set mvarDataBase = Nothing
    
End Property


Public Property Let Recordset(ByVal vData As String)
    Dim sMsg As String
    Dim sTitle As String
    
    '*** Recordset that needs to be displayed
    
    On Error GoTo DbOpenError
    Set mvarRecordset = mvarDataBase.OpenRecordset(vData)
    
    '*** Determine maximum number of records in Recordset
    If mvarRecordset.EOF Then
        mvarMaxRecords = 0
    Else
        mvarRecordset.MoveLast
        mvarMaxRecords = mvarRecordset.RecordCount
        mvarRecordset.MoveFirst
    End If
    
    '*** set up and display scroll bar and progress meter
    '*** progress meter used to indicate how far through table you are
    mvarScrollBar.Max = mvarMaxRecords
    mvarProgressMeter.Max = mvarMaxRecords
    ' if recordset less than number of display rows, disable scroll bar and progress meter
    If mvarMaxRecords < mvarRows Then
        mvarScrollBar.Visible = False
        mvarProgressMeter.Visible = False
    End If
    
    '*** Set up and display grid
    InitalArrayLoad
    DisplayGrid
    Exit Property
    
DbOpenError:
    sTitle = "Grid Class Recordset Error"
    sMsg = Err.Number & " Occurred." & vbCr & Err.Description
    MsgBox Prompt:=sMsg, Buttons:=vbOKOnly + vbCritical, Title:=sTitle
    Set mvarRecordset = Nothing
    
End Property

Public Property Set ScrollBar(ByVal vData As Object)
    Set mvarScrollBar = vData
End Property

Public Property Set ProgressDisplay(ByVal vData As Object)
    Set mvarProgressMeter = vData
End Property

Public Property Let Cols(ByVal vData As Integer)
    mvarCols = vData - 1
    
    ReDim mvarFldNames(mvarCols)
End Property
Public Property Let Rows(ByVal vData As Integer)
    '*** the Columns property must be set first
    
    mvarRows = vData - 1
    
    ReDim mvarTextBoxes(mvarRows, mvarCols)
    ReDim mvarDisplay(mvarRows, mvarCols)
End Property
Public Property Set ColumnDisplay(ByVal vData As TextBox)
    If mvarCol > UBound(mvarTextBoxes) Then
        '**** Error message ??
        Exit Property
    End If
    Set mvarTextBoxes(mvarRow, mvarCol) = vData
End Property

Public Property Let ColumnDataField(ByVal vData As String)
    If mvarCol > UBound(mvarFldNames) Then
        '**** Error message ??
        Exit Property
    End If
    mvarFldNames(mvarCol) = vData
End Property

Public Property Let Col(ByVal vData As Integer)
    mvarCol = vData
End Property
Public Property Let Row(ByVal vData As Integer)
    mvarRow = vData
End Property
Private Sub ClearDisplayArray()

    Dim iRows As Integer
    Dim iCols As Integer
    
    For iRows = 0 To mvarRows
        For iCols = 0 To mvarCols
            mvarDisplay(iRows, iCols) = ""
        Next iCols
    Next iRows
End Sub

Private Sub InitalArrayLoad()

    Dim iRow As Integer
    Dim iCol As Integer
    
    On Error Resume Next
    mvarRecordset.MoveFirst
    
    ClearDisplayArray
    
    '*** recordset exceeds number of display rows
    If mvarMaxRecords > mvarRows Then
        For iRow = 0 To mvarRows
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MoveNext
        Next iRow
    Else
    '*** recordset does not exceed number of display rows
        For iRow = 0 To mvarMaxRecords
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MoveNext
        Next iRow
    End If
    
    '--- set recordset to begining
    mvarRecordset.MoveFirst
    mvarRecPosition = 1
    On Error GoTo 0

End Sub
Private Sub ClearDisplayGrid()

    Dim iRow As Integer
    Dim iCol As Integer
        
    On Error Resume Next
    For iRow = 0 To mvarRows
        For iCol = 0 To mvarCols
            mvarTextBoxes(iRow, iCol) = ""
        Next iCol
    Next iRow
    On Error GoTo 0
    
End Sub

Private Sub DisplayGrid()

    Dim iRow As Integer
    Dim iCol As Integer
        
    On Error Resume Next
    ClearDisplayGrid
    For iRow = 0 To mvarRows
        For iCol = 0 To mvarCols
            mvarTextBoxes(iRow, iCol) = mvarDisplay(iRow, iCol)
        Next iCol
    Next iRow
    On Error GoTo 0
    
End Sub

Public Sub ScrollChange()

'*** this routine will move to the current record as selected by the
'*** scroll bar.   If reach end of recordset, then display last x records.

    Dim iNoOfRecords As Integer
    
    On Error Resume Next
    '*** determine number of records to move, this could be forward or backwards
    If mvarRecPosition = 1 Then
        mvarRecPosition = 2
        mvarRecordset.move mvarRows + 1
        iNoOfRecords = (mvarRows * -1)
    Else
        iNoOfRecords = mvarScrollBar.Value - mvarRecPosition
        '*** if positive move forward, is negative move backwards, set new record position
        '*** if < 1 then movefirst, if > x then move last
        mvarRecPosition = mvarRecPosition + iNoOfRecords
    End If
    Select Case mvarRecPosition
        Case Is <= 1
            mMoveFirst
        Case Is >= mvarMaxRecords
            mMoveLast
        Case Else
            mMove iNoOfRecords
    End Select
End Sub
Private Sub mMoveLast()

    Dim iRow As Integer
    Dim iCol As Integer

    On Error Resume Next
    mvarRecordset.MoveLast
    mvarRecPosition = mvarMaxRecords
    
    ClearDisplayArray
        
    '*** recordset exceeds number of display rows
    If mvarMaxRecords > mvarRows Then
        For iRow = mvarRows To 0 Step -1
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MovePrevious
        Next iRow
    Else
    '*** recordset does not exceed number of display rows
        For iRow = mvarMaxRecords To 0 Step -1
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MovePrevious
        Next iRow
    End If
    mvarRecordset.MoveLast
    
    '*** display records
    DisplayGrid
    
    '*** update ProgressMeter
    mvarProgressMeter.Value = mvarScrollBar.Value

End Sub
Private Sub mMoveFirst()
    
    Dim iRow As Integer
    Dim iCol As Integer
    
    On Error Resume Next
    mvarRecordset.MoveFirst
    mvarRecPosition = 1
    
    ClearDisplayArray
        
    '*** recordset exceeds number of display rows
    If mvarMaxRecords > mvarRows Then
        For iRow = 0 To mvarRows
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MoveNext
            mvarRecPosition = mvarRecPosition + 1
        Next iRow
    Else
    '*** recordset does not exceed number of display rows
        For iRow = 0 To mvarMaxRecords
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MoveNext
            mvarRecPosition = mvarRecPosition + 1
        Next iRow
    End If
    
    '*** display records
    DisplayGrid
    
    '*** update ProgressMeter
    mvarProgressMeter.Value = mvarScrollBar.Value

End Sub
Private Sub mMove(ByVal vData As Integer)
    
    Dim iRow As Integer
    Dim iCol As Integer
    
    On Error Resume Next
    
    '*** if records is in last x of dataset, need to reset where moving from
    If mvarRecPosition > (mvarMaxRecords - mvarRows) Then
        mvarRecordset.move (vData - mvarRows)
    Else
        mvarRecordset.move vData
    End If
    If mvarRecordset.EOF Then
        mMoveLast
        Exit Sub
    ElseIf mvarRecordset.BOF Then
        mMoveFirst
        Exit Sub
    End If
    
    ClearDisplayArray
        
    '*** recordset exceeds number of display rows
    If mvarMaxRecords > mvarRows Then
        For iRow = 0 To mvarRows
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MoveNext
            mvarRecPosition = mvarRecPosition + 1
        Next iRow
    Else
    '*** recordset does not exceed number of display rows
        For iRow = 0 To mvarMaxRecords
            For iCol = 0 To mvarCols
                mvarDisplay(iRow, iCol) = mvarRecordset.Fields(mvarFldNames(iCol))
            Next iCol
            mvarRecordset.MoveNext
            mvarRecPosition = mvarRecPosition + 1
        Next iRow
    End If
    
    '*** display records
    DisplayGrid
    
    '*** update ProgressMeter
    mvarProgressMeter.Value = mvarScrollBar.Value

End Sub

Public Sub MoveFirst()
    mMoveFirst
    mvarScrollBar.Value = 1
End Sub
Public Sub MoveLast()
    mMoveLast
    mvarScrollBar.Value = mvarMaxRecords
End Sub
Public Sub MovePrevious()
    mvarRecPosition = mvarRecPosition - (mvarRows + 2)
    mMove ((mvarRows + 2) * -1)
    mvarScrollBar.Value = mvarRecPosition - (mvarRows + 1)
End Sub
Public Sub MoveNext()
    If mvarRecPosition = 1 Then
        mvarRecPosition = 2
        mvarRecordset.move mvarRows + 1
    Else
        mvarRecPosition = mvarRecPosition - (mvarRows)
    End If
    mMove ((mvarRows) * -1)
    mvarScrollBar.Value = mvarRecPosition - (mvarRows + 1)
End Sub

Public Sub move(ByVal vData As Integer)
    mvarRecPosition = mvarRecPosition + vData
    mMove (vData - (mvarRows + 1))
    mvarScrollBar.Value = mvarRecPosition - (mvarRows + 1)
End Sub
