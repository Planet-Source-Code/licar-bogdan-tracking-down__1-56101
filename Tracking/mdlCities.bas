Attribute VB_Name = "mdlCities"
Option Explicit

Public Type City    'The user defined type
    Name As String
    X As Integer
    Y As Integer
End Type

'These are the cities that I chose to represent on the map
Public Rome As City
Public New_York As City
Public Sidney As City
Public Moscow As City
Public Cape_Town As City
Public Rio As City
Public San_Francisco As City
Public Tokyo As City
Public Vladivostok As City
Public Chicago As City
Public Hong_Kong As City
Public Paris As City
Public Prague As City
Public Bombay As City
Public Oslo As City
Public Buenos_Aires As City

Public SelCity As City

'The number of the cities; if you add some new cites just change this constant
Public Const Cities_N° = 16

'All cities and their coordinates on a map with x = 1000, y = 1000
Public Sub Cities()
Moscow.Name = "Moscow"
Moscow.X = 597
Moscow.Y = 761

Rome.Name = "Rome"
Rome.X = 511
Rome.Y = 682

New_York.Name = "New York"
New_York.X = 269
New_York.Y = 679

Sidney.Name = "Sidney"
Sidney.X = 907
Sidney.Y = 360

Cape_Town.Name = "Cape Town"
Cape_Town.X = 542
Cape_Town.Y = 362

Rio.Name = "Rio de Janeiro"
Rio.X = 369
Rio.Y = 461

San_Francisco.Name = "San Francisco"
San_Francisco.X = 124
San_Francisco.Y = 654

Tokyo.Name = "Tokyo"
Tokyo.X = 877
Tokyo.Y = 658

Vladivostok.Name = "Vladivostok"
Vladivostok.X = 926
Vladivostok.Y = 731

Chicago.Name = "Chicago"
Chicago.X = 223
Chicago.Y = 694

Hong_Kong.Name = "Hong Kong"
Hong_Kong.X = 813
Hong_Kong.Y = 600

Paris.Name = "Paris"
Paris.X = 480
Paris.Y = 709

Prague.Name = "Prague"
Prague.X = 521
Prague.Y = 714

Bombay.Name = "Bombay"
Bombay.X = 675
Bombay.Y = 598

Oslo.Name = "Oslo"
Oslo.X = 503
Oslo.Y = 772

Buenos_Aires.Name = "Buenos Aires"
Buenos_Aires.X = 292
Buenos_Aires.Y = 367

End Sub


