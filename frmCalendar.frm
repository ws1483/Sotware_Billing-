VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Pick a date"
   ClientHeight    =   5328
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7764
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================================
' frmCalendar — pure-VBA popup month calendar (no ActiveX).
'   Design-time controls required on the form:
'     cboMonth   (ComboBox)
'     cboYear    (ComboBox)
'     btnD1      (CommandButton)  <- top-left day cell; other 41 auto-created
'     btnToday   (CommandButton)  Caption "Today"
'     btnClear   (CommandButton)  Caption "Clear"
'   Optional: 7 day-name labels (Mo Tu We Th Fr Sa Su) placed above btnD1.
'
'   Usage (from a module):  d = PickDate([seed])
'     -> returns a Date, or Empty if Cleared/Cancelled.
' ============================================================================

Public SelectedDate As Variant     ' Date, or Empty if cleared/cancelled
Public Cancelled As Boolean

Private mDayBtns(1 To 42) As MSForms.CommandButton
Private mHandlers As Collection     ' holds clsDayBtn instances (keep them alive)
Private mCurYear As Integer
Private mCurMonth As Integer

' ---------------------------------------------------------------- INIT ------
Private Sub UserForm_Initialize()
    Dim i As Integer, y As Integer

    Cancelled = True
    SelectedDate = Empty

    ' month names
    cboMonth.Clear
    For i = 1 To 12
        cboMonth.AddItem Format(DateSerial(2000, i, 1), "mmmm")
    Next i

    ' year range (current -5 .. +5)
    cboYear.Clear
    For y = Year(Date) - 5 To Year(Date) + 5
        cboYear.AddItem y
    Next y

    BuildGrid

    mCurYear = Year(Date)
    mCurMonth = Month(Date)
    cboMonth.ListIndex = mCurMonth - 1
    cboYear.Value = mCurYear

    RefreshGrid
End Sub

' Public: allow caller to open on a specific date's month
Public Sub SeedDate(ByVal seed As Variant)
    If IsDate(seed) Then
        mCurYear = Year(seed)
        mCurMonth = Month(seed)
        cboMonth.ListIndex = mCurMonth - 1
        cboYear.Value = mCurYear
        RefreshGrid
    End If
End Sub

' ----------------------------------------------------- BUILD DAY GRID -------
Private Sub BuildGrid()
    Dim i As Integer, baseL As Single, baseT As Single, w As Single, h As Single
    Dim col As Integer, row As Integer
    Dim hnd As clsDayBtn

    ' template = btnD1
    Set mDayBtns(1) = Me.btnD1
    baseL = Me.btnD1.Left
    baseT = Me.btnD1.Top
    w = Me.btnD1.Width
    h = Me.btnD1.Height

    ' create the other 41
    For i = 2 To 42
        Set mDayBtns(i) = Me.Controls.Add("Forms.CommandButton.1", "btnD" & i, True)
    Next i

    ' position all 42 in a 6-row x 7-col grid
    For i = 1 To 42
        col = (i - 1) Mod 7
        row = (i - 1) \ 7
        With mDayBtns(i)
            .Left = baseL + col * w
            .Top = baseT + row * h
            .Width = w
            .Height = h
            .Font.Size = 9
            .TakeFocusOnClick = False
        End With
    Next i

    ' hook click events for every day button
    Set mHandlers = New Collection
    For i = 1 To 42
        Set hnd = New clsDayBtn
        Set hnd.Btn = mDayBtns(i)
        Set hnd.Frm = Me
        mHandlers.Add hnd
    Next i
End Sub

' --------------------------------------------------- REFRESH GRID ----------
Private Sub RefreshGrid()
    Dim dim1 As Date, firstDow As Integer, daysInMonth As Integer
    Dim i As Integer, dNum As Integer

    If mCurMonth < 1 Or mCurMonth > 12 Then Exit Sub

    dim1 = DateSerial(mCurYear, mCurMonth, 1)
    firstDow = Weekday(dim1, vbMonday)                       ' Mon=1 .. Sun=7
    daysInMonth = Day(DateSerial(mCurYear, mCurMonth + 1, 0))

    Me.Caption = "Pick a date — " & Format(dim1, "mmmm yyyy")

    For i = 1 To 42
        dNum = i - firstDow + 1
        If dNum >= 1 And dNum <= daysInMonth Then
            mDayBtns(i).Caption = CStr(dNum)
            mDayBtns(i).Enabled = True
            ' highlight today
            If DateSerial(mCurYear, mCurMonth, dNum) = Date Then
                mDayBtns(i).BackColor = RGB(200, 230, 255)
            Else
                mDayBtns(i).BackColor = &H8000000F            ' default button face
            End If
        Else
            mDayBtns(i).Caption = ""
            mDayBtns(i).Enabled = False
            mDayBtns(i).BackColor = &H8000000F
        End If
    Next i
End Sub

' --------------------------------------------------- DROPDOWN EVENTS -------
Private Sub cboMonth_Change()
    If cboMonth.ListIndex >= 0 Then
        mCurMonth = cboMonth.ListIndex + 1
        RefreshGrid
    End If
End Sub

Private Sub cboYear_Change()
    If IsNumeric(cboYear.Value) Then
        mCurYear = CInt(cboYear.Value)
        RefreshGrid
    End If
End Sub

' --------------------------------------------------- PICK / BUTTONS --------
' Called by clsDayBtn when a day button is clicked
Public Sub PickDay(ByVal dNum As Integer)
    If dNum < 1 Then Exit Sub
    SelectedDate = DateSerial(mCurYear, mCurMonth, dNum)
    Cancelled = False
    Me.Hide
End Sub

Private Sub btnToday_Click()
    SelectedDate = Date
    Cancelled = False
    Me.Hide
End Sub

Private Sub btnClear_Click()
    SelectedDate = Empty
    Cancelled = False        ' user chose to clear (not a cancel)
    Me.Hide
End Sub

' Treat the [X] window close as a cancel
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        SelectedDate = Empty
        ' allow it to close normally
    End If
End Sub

