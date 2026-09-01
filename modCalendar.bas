Attribute VB_Name = "modCalendar"
Option Explicit
' Returns picked Date, or Empty if cancelled/cleared.
Public Function PickDate(Optional ByVal seed As Variant) As Variant
    Dim f As frmCalendar
    Set f = New frmCalendar
    If IsDate(seed) Then f.SeedDate seed     ' open on the cell's current month
    f.Show                                    ' modal
    If f.Cancelled Then
        PickDate = Empty
    Else
        PickDate = f.SelectedDate             ' Date, or Empty if Clear was pressed
    End If
    Unload f
End Function

