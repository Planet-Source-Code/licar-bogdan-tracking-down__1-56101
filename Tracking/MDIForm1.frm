VERSION 5.00
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Tracking You Down Across The World"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8835
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu default_pos 
         Caption         =   "&Default Windows Positions"
      End
      Begin VB.Menu g 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu load 
      Caption         =   "&Load New Map"
      Begin VB.Menu physical 
         Caption         =   "&Physical Map"
      End
      Begin VB.Menu day 
         Caption         =   "&Earth By Day"
      End
      Begin VB.Menu night 
         Caption         =   "&Earth By Night"
      End
   End
   Begin VB.Menu windows 
      Caption         =   "&Windows"
      Begin VB.Menu action 
         Caption         =   "&Action"
         Checked         =   -1  'True
      End
      Begin VB.Menu order 
         Caption         =   "&Choose Your Order"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub action_Click()
If action.Checked = True Then
action.Checked = False
frmCommands.Hide
ElseIf action.Checked = False Then
action.Checked = True
frmCommands.Show
End If
End Sub

Private Sub default_pos_Click()
Default_Positions
End Sub

Private Sub exit_Click()
Unload Me
End
End Sub

Private Sub day_Click()
frmMap.Picture1.Picture = LoadPicture(App.Path + "\Maps\Earth By Day.jpg")
frmMap.Caption = " Map - Earth By Day"
End Sub

Private Sub MDIForm_Load()
frmMap.Enabled = True
frmSel.Enabled = True
frmCommands.Enabled = True
End Sub

Private Sub night_Click()
frmMap.Picture1.Picture = LoadPicture(App.Path + "\Maps\Earth By Night.jpg")
frmMap.Caption = " Map - Earth By Night"
End Sub

Private Sub order_Click()
If order.Checked = True Then
order.Checked = False
frmSel.Hide
ElseIf order.Checked = False Then
order.Checked = True
frmSel.Show
End If
End Sub

Private Sub physical_Click()
frmMap.Picture1.Picture = LoadPicture(App.Path + "\Maps\Physical Map.jpg")
frmMap.Caption = " Map - Physical"
End Sub


