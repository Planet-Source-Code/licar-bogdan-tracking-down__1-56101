VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCommands 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Action"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1530
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   1530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Colors"
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Dots"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   795
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   240
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Lines"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   315
         Width           =   855
      End
      Begin VB.Shape Shape4 
         Height          =   375
         Left            =   240
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3960
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   420
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Width           =   1575
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Text            =   "1"
         Top             =   1365
         Width           =   485
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H80000007&
         Caption         =   "Start"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   180
         Width           =   975
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Value           =   10
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   2280
         Width           =   615
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1560
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   1560
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1560
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   1560
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   1560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Seconds To Remain In A City:"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Line Speed:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon2 
      Left            =   480
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   960
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim col1 As Long, col2 As Long, col As Long, col_lines As Long
Dim xp As Double, yp As Double, st As Double
Dim Curr_City As City
Dim Next_City As City
Dim interv As Double


Private Sub cmdStart_Click()
interv = CDbl(Text2.Text)
interv = interv * 100
Timer1.Interval = interv    'The default interval is 0,1 second

frmMap.StatusBar1.Panels(2).Text = Text2.Text & " Second(s) In A City"
frmMap.StatusBar1.Panels(3).Text = "Speed: " & Text1.Text
a = 0
frmMap.ProB.Max = Cities_N° + 1 'Just sets the maximum value of the ProgressBar

If cmdStart.Caption = "Start" Then
cmdStart.Caption = "Stop"
Timer1.Enabled = True

ElseIf cmdStart.Caption = "Stop" Then
cmdStart.Caption = "Start"
Timer1.Enabled = False
End If

End Sub

Private Sub cmdReset_Click()
frmMap.Picture1.Cls
a = 0   'Resets the counter's value, used in the timer
frmMap.StatusBar1.Panels.Item(1).Text = ""
frmMap.ProB.Value = 0
End Sub

Private Sub Form_Load()
With frmCommands
    .Left = 11320
    .Height = 4335
    .Top = 100
End With

Text1.Text = VScroll1.Value

SetTheScale frmMap.Picture1, 0, 1000, 1000, 0   'Sets the value of the 2 axis
Cities
col2 = vbRed
End Sub

Function SetTheScale(ByVal obj As Object, ByVal upper_left_x As Single, _
ByVal upper_left_y As Single, ByVal lower_right_x As Single, ByVal lower_right_y As Single)
    obj.ScaleLeft = upper_left_x
    obj.ScaleTop = upper_left_y
    obj.ScaleWidth = lower_right_x - upper_left_x
    obj.ScaleHeight = lower_right_y - upper_left_y
End Function

Private Sub Form_Unload(Cancel As Integer)
MDI.action.Checked = False
End Sub

Private Sub Label3_Click()
    On Error Resume Next            'To animate the point are used 2 colors
    dlgCommon.Flags = cdlCCRGBInit
    dlgCommon.ShowColor
    col1 = dlgCommon.Color
    col = col1
    
    dlgCommon2.Flags = cdlCCRGBInit
    dlgCommon2.ShowColor
    col2 = dlgCommon2.Color
End Sub

Private Sub Label7_Click()
    On Error Resume Next
    dlgCommon.Flags = cdlCCRGBInit
    dlgCommon.ShowColor
    col_lines = dlgCommon.Color
End Sub

Private Sub Timer1_Timer()
a = a + 0.1     'Starts the counter; if the timer's interval is set to 0,1 (default)
                'then it will remain in a city 1 second due to the a = a + 0,1
                
frmMap.ProB.Value = a   'Starts the progressbar
If a = 0.1 Then First   'Goes directly to the first chosen city

'When the counter hits an integer value it goes to the other chosen cities
'from the ComboBoxes.
If a = 2 Or a = 3 Or a = 4 Or a = 5 Or a = 6 Or a = 7 Or a = 8 Or a = 9 _
Or a = 10 Or a = 11 Or a = 12 Or a = 13 Or a = 14 Or a = 15 Then Lines: The_Others

If a = Cities_N° Then Lines: Last  'When the counter has reached the number of
                                   'cities it goes to the last ComboBox

If a = frmMap.ProB.Max Then cmdStart_Click 'Stops the timer

Points 'It draws the points waiting for the line to go into the next city

frmMap.StatusBar1.Panels.Item(1).Text = Curr_City.Name
End Sub

Private Sub Lines()
    Dim sp As Double
    sp = Text1.Text * 0.001 'The speed of the lines
    
    
    If Curr_City.X < Next_City.X Then st = sp   'It assigns a positive or a negative
    If Curr_City.X > Next_City.X Then st = -sp  'value to the step comparing the x
                                                'of the current and next city
    
    frmMap.Picture1.DrawWidth = 2
    
    For xp = Curr_City.X To Next_City.X Step st
    
    'Finds the equation of the points
    yp = Coef_X(Curr_City.X, Curr_City.Y, Next_City.X, Next_City.Y) * xp + _
         Ord_Y(Curr_City.X, Curr_City.Y, Next_City.X, Next_City.Y)
    
    frmMap.Picture1.PSet (xp, yp), col_lines 'Just draws the points
    Next xp
    
End Sub

Private Sub Points()
frmMap.Picture1.DrawWidth = 10

If col = col1 Then  'Makes the points blink
    col = col2
Else
    col = col1
End If

frmMap.Picture1.PSet (Curr_City.X, Curr_City.Y), col
End Sub

Private Sub First()
a = 0   'Assigns 0 to the counter and goes to the Select Case in the module
The_Index
Curr_City = SelCity 'The current city is the one in the first ComboBox

a = 1   'Assigns 1 so that the 2nd ComboBox can be changed
The_Index
Next_City = SelCity 'The next city is the one in the 2nd ComboBox
a = 1.1
End Sub

Private Sub The_Others()
The_Index
Curr_City = Next_City   'The next city becomes the current city
Next_City = SelCity     'The next city will automatically be the one into the
                        'next ComboBox
End Sub

Private Sub Last()
Curr_City = Next_City
End Sub

Private Sub VScroll1_Change()
Text1.Text = VScroll1.Value
Text1.SetFocus
End Sub


