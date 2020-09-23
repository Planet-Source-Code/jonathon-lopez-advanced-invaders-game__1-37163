VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Fight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myshots(1 To 21) As shot
Private myShips(1 To 2) As Star
Private dead(1 To 2) As Integer
Private numShots As Integer
Private Ticker As Integer, rShot As Integer, bShot As Integer
Private init As Boolean
Private sShot(1 To 2) As Integer

Public Sub CreateFight()
Dim i As Integer
dead(1) = 0
dead(2) = 0
Ticker = 0
init = True
Randomize
rShot = Int(Rnd * 10 + 1)
bShot = Int(Rnd * 10 + 1)
i = Int(Rnd * 480)
myShips(1).Y = i
i = i + Int(Rnd * 40) - 20
If i < 0 Then i = 0
If i > 480 Then i = 480
myShips(2).Y = i
numShots = Int(Rnd * 2 + 1)
myShips(numShots).X = 0
If numShots = 1 Then
myShips(2).X = 640
Else
myShips(1).X = 640
End If
myShips(1).vx = Int(Rnd * 3 + 1)
myShips(2).vx = -Int(Rnd * 3 + 1)
myShips(1).vy = 0
myShips(2).vy = 0
sShot(1) = -1
sShot(2) = -1
numShots = 0
End Sub

Public Sub DoFight()

Dim i As Integer, j As Integer, dx As Integer, dy As Integer, d As Integer

myShips(1).X = myShips(1).X + myShips(1).vx
myShips(1).Y = myShips(1).Y + myShips(1).vy
myShips(2).X = myShips(2).X + myShips(2).vx
myShips(2).Y = myShips(2).Y + myShips(2).vy


If myShips(1).X > myShips(2).X Then myShips(1).vx = myShips(1).vx - 1

If myShips(1).X < myShips(2).X Then myShips(1).vx = myShips(1).vx + 1

If myShips(1).Y > myShips(2).Y Then myShips(1).vy = myShips(1).vy - 1

If myShips(1).Y < myShips(2).Y Then myShips(1).vy = myShips(1).vy + 1


If myShips(2).X > myShips(1).X Then myShips(2).vx = myShips(2).vx - 1

If myShips(2).X < myShips(1).X Then myShips(2).vx = myShips(2).vx + 1

If myShips(2).Y > myShips(1).Y Then myShips(2).vy = myShips(2).vy - 1

If myShips(2).Y < myShips(1).Y Then myShips(2).vy = myShips(2).vy + 1

If myShips(1).vx > 5 Then myShips(1).vx = 5
If myShips(2).vx > 5 Then myShips(2).vx = 5
If myShips(1).vx < -5 Then myShips(1).vx = -5
If myShips(2).vx < -5 Then myShips(2).vx = -5

If myShips(1).vy > 5 Then myShips(1).vy = 5
If myShips(2).vy > 5 Then myShips(2).vy = 5
If myShips(1).vy < -5 Then myShips(1).vy = -5
If myShips(2).vy < -5 Then myShips(2).vy = -5

For i = 1 To 2

    If myShips(i).X < 0 Then
        myShips(i).vx = -myShips(i).vx
        myShips(i).X = -myShips(i).X
    End If

    If myShips(i).X > 640 Then
        myShips(i).vx = -myShips(i).vx
        myShips(i).X = 635
    End If
    
    If myShips(i).Y < 0 Then
        myShips(i).vy = -myShips(i).vy
        myShips(i).Y = -myShips(i).Y
    End If

    If myShips(i).Y > 480 Then
        myShips(i).vy = -myShips(i).vy
        myShips(i).Y = 475
    End If
    
Next i

Ticker = Ticker + 1

If rShot = Ticker And numShots < 20 Then
    numShots = numShots + 1
    myshots(numShots).X = myShips(1).X + 2
    myshots(numShots).Y = myShips(1).Y + 1
    myshots(numShots).vx = -(myShips(1).X - myShips(2).X) / 20
    myshots(numShots).vy = -(myShips(1).Y - myShips(2).Y) / 20
    Randomize
    i = Int(Rnd * 20)
    If i = 1 Then
        myshots(numShots).t = 1
    Else
        myshots(numShots).t = 0
    End If
    If myshots(numShots).vx = 0 And myshots(numShots).vy = 0 Then
        myshots(numShots).vx = 1
        myshots(numShots).vy = 1
    End If
End If

If bShot = Ticker And numShots < 20 Then
    numShots = numShots + 1
    myshots(numShots).X = myShips(1).X + 2
    myshots(numShots).Y = myShips(1).Y + 1
    myshots(numShots).vx = (myShips(1).X - myShips(2).X) / 10
    myshots(numShots).vy = (myShips(1).Y - myShips(2).Y) / 10
    Randomize
    i = Int(Rnd * 20)
    If i = 1 Then
        myshots(numShots).t = 1
    Else
        myshots(numShots).t = 0
    End If
    If myshots(numShots).vx = 0 And myshots(numShots).vy = 0 Then
        myshots(numShots).vx = 1
        myshots(numShots).vy = 1
    End If
End If

If Ticker > 10 Then Ticker = 0

For i = 1 To numShots
    myshots(i).X = myshots(i).X + myshots(i).vx
    myshots(i).Y = myshots(i).Y + myshots(i).vy
    If myshots(i).X > 640 Or myshots(i).X < 0 Or myshots(i).Y > 480 Or myshots(i).Y < 0 Then
        For j = i To numShots
            myshots(j).X = myshots(j + 1).X
            myshots(j).Y = myshots(j + 1).Y
            myshots(j).vx = myshots(j + 1).vx
            myshots(j).vy = myshots(j + 1).vy
        Next j
        numShots = numShots - 1
    End If
Next i

For j = 1 To 2
    If dead(j) = 1 Then
        myShips(j).X = (Int(Rnd * 2)) * 640
        myShips(j).Y = Int(Rnd * 470)
        dead(j) = 0
    End If
Next j

For i = 1 To numShots
    For j = 1 To 2
        If (myshots(i).X - myShips(j).X) > 0 And (myshots(i).X - myShips(j).X) < 11 And (myshots(i).Y - myShips(j).Y) > 0 And (myshots(i).Y - myShips(j).Y) < 4 Then
            dead(j) = 1
        End If
    Next j
Next i

End Sub

Property Get isInit() As Boolean

isInit = init

End Property

Property Let isInit(ini As Boolean)

init = ini

End Property

Property Get shipx(i As Integer) As Integer
shipx = myShips(i).X
End Property
Property Get shipy(i As Integer) As Integer
shipy = myShips(i).Y
End Property

Property Get nShots() As Integer
nShots = numShots
End Property

Property Get shotX(i As Integer) As Integer
shotX = myshots(i).X
End Property
Property Get shotY(i As Integer) As Integer
shotY = myshots(i).Y
End Property
Property Get shotT(i As Integer) As Integer
shotT = myshots(i).t
End Property

Property Get Death(i As Integer) As Integer
Death = dead(i)
End Property
