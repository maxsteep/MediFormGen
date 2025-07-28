VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
     'check cells for desired format to trigger the calendarfrm.show routine
     'otherwise exit the sub
    Dim DateFormats, DF
    DateFormats = Array("m/d/yy;@", "mmmm d yyyy")
    For Each DF In DateFormats
        If DF = Target.NumberFormat Then
            If CalendarFrm.HelpLabel.Caption <> "" Then
                CalendarFrm.Height = 191 + CalendarFrm.HelpLabel.Height
            Else: CalendarFrm.Height = 191
                CalendarFrm.Show
            End If
        End If
    Next
End Sub
 



