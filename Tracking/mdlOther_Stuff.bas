Attribute VB_Name = "mdlOther_Stuff"
Public i As Integer
Public a As Currency

'These two functions are for drawing the lines; yes I could have used the "line()-()"
'but I wouldn't have the lines drawn point by point
Public Function Coef_X(x1, y1, x2, y2)
Dim xx, yy, int1, int2, int3, int4, Coef As Double

'It basically finds the coefficient of a line given the coordinates of two points

xx = 1
yy = 1

int1 = yy * (x2 - x1)
int2 = -y1 * (x2 - x1)
int3 = xx * (y2 - y1)
int4 = -x1 * (y2 - y1)
On Error Resume Next
Coef = int3 / int1
Coef_X = Coef

End Function

Public Function Ord_Y(x1, y1, x2, y2)
Dim xx, yy, int1, int2, int3, int4, Ord As Double

'It finds the point (0, ord) where the line intersects y axis

xx = 1
yy = 1

int1 = yy * (x2 - x1)
int2 = -y1 * (x2 - x1)
int3 = xx * (y2 - y1)
int4 = -x1 * (y2 - y1)
On Error Resume Next
Ord = (int4 - int2) / int1
Ord_Y = Ord

End Function

Public Sub Default_Positions()
With frmSel             'It restores the default positions to all forms
    .Left = 12800
    .Width = 2500
    .Height = 6735
    .Top = 100
End With
With frmCommands
    .Left = 11320
    .Height = 4335
    .Top = 100
End With
With frmMap
    .Height = 9025
    .Width = 11220
    .Top = 0
    .Left = 0
End With

End Sub

Public Sub The_Index()
Dim Town As String

Town = frmSel.Combo1(a).Text

'When the user chooses his/her own order this checks what city has been chosen and
'gives to the Selected city the chosen city's properties

Select Case Town
    Case "Bombay"
    SelCity = Bombay
    
    Case "Buenos Aires"
    SelCity = Buenos_Aires
    
    Case "Cape Town"
    SelCity = Cape_Town
    
    Case "Chicago"
    SelCity = Chicago
    
    Case "Hong Kong"
    SelCity = Hong_Kong
    
    Case "Moscow"
    SelCity = Moscow
    
    Case "New York"
    SelCity = New_York
    
    Case "Oslo"
    SelCity = Oslo
    
    Case "Paris"
    SelCity = Paris
    
    Case "Prague"
    SelCity = Prague
    
    Case "Rio"
    SelCity = Rio
    
    Case "Rome"
    SelCity = Rome
    
    Case "San Francisco"
    SelCity = San_Francisco
    
    Case "Sidney"
    SelCity = Sidney
    
    Case "Tokyo"
    SelCity = Tokyo
    
    Case "Vladivostok"
    SelCity = Vladivostok
End Select
End Sub

