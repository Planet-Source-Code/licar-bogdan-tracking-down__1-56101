VERSION 5.00
Begin VB.Form frmSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Choose Your Order"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2460
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Random"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   5880
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   15
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5520
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   14
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5160
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   13
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   12
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4440
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   11
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   10
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   9
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   8
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   7
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   6
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   5
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label16 
      Caption         =   "16°"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "15°"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "14°"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "13°"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "12°"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "11*"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "10°"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "9°"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "8°"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "7°"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "6°"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "5°"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "4°"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "3°"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "2°"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "1°"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ran As Integer, j As Integer
Dim pr(0 To (Cities_N° - 1)) As Integer

'Gives you a random order of the cities into the ComboBoxes
Private Sub Command1_Click()

For i = 0 To Cities_N° - 1  'Assigns to each element of the array -1;
    pr(i) = -1              'it could be also -2314 or 524; just a value <> 0 to 15
Next i

Randomize
For i = 0 To Cities_N° - 1

Back:
ran = Rnd * (Cities_N° - 1) 'A random number from 0 to 15
    
    For j = 0 To i - 1                  'This is a verifying loop. It checks if the
        If ran = pr(j) Then GoTo Back   'random value has been already given.
    Next j                              'In that case turns back until a new value
                                        'is assigned.

Combo1(i).Text = Combo1(i).List(ran)    'Gives to the ComboBox the random item

pr(i) = ran 'Puts the random value into the array

Next i
End Sub

Private Sub Form_Load()
With frmSel
    .Left = 12800
    .Width = 2500
    .Height = 6735
    .Top = 100
End With

'Adds to each ComboBox every city
For i = 0 To Cities_N° - 1
Combo1(i).AddItem "Sidney"
Combo1(i).AddItem "Moscow"
Combo1(i).AddItem "New York"
Combo1(i).AddItem "Prague"
Combo1(i).AddItem "San Francisco"
Combo1(i).AddItem "Chicago"
Combo1(i).AddItem "Buenos Aires"
Combo1(i).AddItem "Rio"
Combo1(i).AddItem "Rome"
Combo1(i).AddItem "Cape Town"
Combo1(i).AddItem "Vladivostok"
Combo1(i).AddItem "Bombay"
Combo1(i).AddItem "Oslo"
Combo1(i).AddItem "Paris"
Combo1(i).AddItem "Hong Kong"
Combo1(i).AddItem "Tokyo"
Combo1(i).Text = Combo1(i).List(i)
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDI.order.Checked = False
End Sub
