VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9600
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox shippic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   8280
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   390
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox pal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   9120
      Picture         =   "frmMain.frx":06D2
      ScaleHeight     =   15
      ScaleWidth      =   90
      TabIndex        =   13
      Top             =   6240
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox eShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   8640
      Picture         =   "frmMain.frx":0728
      ScaleHeight     =   315
      ScaleWidth      =   390
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox eShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   8160
      Picture         =   "frmMain.frx":0DFA
      ScaleHeight     =   315
      ScaleWidth      =   390
      TabIndex        =   11
      Top             =   6360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox eShot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   7320
      Picture         =   "frmMain.frx":14CC
      ScaleHeight     =   120
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox eShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   7680
      Picture         =   "frmMain.frx":168E
      ScaleHeight     =   315
      ScaleWidth      =   390
      TabIndex        =   9
      Top             =   6360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox shot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   7320
      Picture         =   "frmMain.frx":1D60
      ScaleHeight     =   120
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox ship 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7800
      Picture         =   "frmMain.frx":1F22
      ScaleHeight     =   315
      ScaleWidth      =   390
      TabIndex        =   7
      Top             =   6840
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox uShot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   30
      Index           =   1
      Left            =   8280
      Picture         =   "frmMain.frx":25F4
      ScaleHeight     =   30
      ScaleWidth      =   120
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox uRed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   1
      Left            =   8760
      Picture         =   "frmMain.frx":2666
      ScaleHeight     =   45
      ScaleWidth      =   165
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox uBlue 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   1
      Left            =   8520
      Picture         =   "frmMain.frx":2714
      ScaleHeight     =   45
      ScaleWidth      =   165
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox uShot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   8280
      Picture         =   "frmMain.frx":27C2
      ScaleHeight     =   15
      ScaleWidth      =   30
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox uBlue 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   0
      Left            =   8520
      Picture         =   "frmMain.frx":280C
      ScaleHeight     =   45
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox uRed 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   0
      Left            =   8760
      Picture         =   "frmMain.frx":28BA
      ScaleHeight     =   45
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox buff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   9240
      ScaleHeight     =   7155
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   9615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_DblClick()
End
End Sub
Private Sub Form_Load()
Dim sldwn As Long, sCounter As Long
Dim genStar As Integer, i As Integer, j As Integer, genbHole As Integer, genFight As Integer
Dim myFight(1 To 10) As New cls_Fight, tStar As Star, aShoot As Integer, atShoot As Integer
Dim tsc As Integer, tX As Integer, tY As Integer, k As Integer, aShot As Star, aShip As Star
Dim sec As Integer, frame As Integer, lframe As Integer, l As Integer, m As Integer, wcount As Integer
For i = 0 To 25
    For j = 0 To 20
        mFire(i, j) = False
    Next j
Next i
For i = 1 To 4
    For j = 1 To 10
        Fire(i, j) = 15
    Next j
Next i
gGame = True

nLevel = 1
ntShots = 0
nFight = 0
aShoot = 0
atShoot = 0
nShots = 0
myShip.X = 307
myShip.Y = 480 - ship.Height
Randomize
For i = 1 To nLevel * 5
    nFight = nFight + 1
    tShips(nFight).X = Int(Rnd * 600)
    tShips(nFight).Y = Int(Rnd * 250)
    tShips(nFight).vx = 0
    tShips(nFight).vy = 0
    tShips(nFight).t = 0
    tShips(nFight).d = 0
Next i

For i = 1 To 10
myFight(i).isInit = False
Next i
hbH = False

sNum = 0
frame = 0
sec = Second(Time)

sldwn = 100000

Show

Do
If sCounter = sldwn Then
    sCounter = 0
    DoEvents
    buff.Cls
    Randomize
        
    If buff.BackColor <> vbBlack Then buff.BackColor = vbBlack
    
    If nFight < 1 And nShots = 0 Then
        nLevel = nLevel + 1
        nFight = nLevel * 5
        For i = 1 To nFight
            Randomize
            tShips(i).X = Int(Rnd * 600)
            tShips(i).Y = Int(Rnd * 250)
            tShips(i).vx = 0
            tShips(i).vy = 0
            tShips(i).t = 0
        Next i
    End If
    
    For j = 1 To 5
    If myFight(j).isInit = False Then
        genFight = Int(Rnd * 10)
        If genFight = 5 Then myFight(j).CreateFight
    Else
        myFight(j).DoFight
        
        '' disable this part if you dont want to see the backgound fights
        
        BitBlt buff.hdc, myFight(j).shipx(1), myFight(j).shipy(1), uRed(0).Width, uRed(0).Height, uRed(myFight(j).Death(1)).hdc, 0, 0, SRCCOPY
        BitBlt buff.hdc, myFight(j).shipx(2), myFight(j).shipy(2), uBlue(0).Width, uBlue(0).Height, uBlue(myFight(j).Death(2)).hdc, 0, 0, SRCCOPY
        For i = 1 To myFight(j).nShots
            BitBlt buff.hdc, myFight(j).shotX(i), myFight(j).shotY(i), 2 + myFight(j).shotT(i) * 4, 1 + myFight(j).shotT(i) * 1, uShot(myFight(j).shotT(i)).hdc, 0, 0, SRCCOPY
        Next i
    End If
    Next j
    If hbH = False Then
        genbHole = Int(Rnd * 100)
        If genbHole = 5 Then
            hbH = True
            bHole.X = Int(Rnd * 640)
            bHole.Y = 0
            bHole.vy = Int(Rnd * 2 + 1)
        End If
    Else
        SetPixel buff.hdc, bHole.X, bHole.Y, &HFF00FF
        SetPixel buff.hdc, bHole.X - 1, bHole.Y, &HFF00FF
        SetPixel buff.hdc, bHole.X + 1, bHole.Y, &HFF00FF
        SetPixel buff.hdc, bHole.X, bHole.Y - 1, &HFF00FF
        SetPixel buff.hdc, bHole.X, bHole.Y + 1, &HFF00FF
        For i = 1 To sNum
            Dim d As Integer, dx As Integer, dy As Integer
            d = ((sArr(i).X - bHole.X) ^ 2 + (sArr(i).Y - bHole.Y) ^ 2) ^ 0.5
            On Error Resume Next
            dx = (sArr(i).X - bHole.X) / d
            dy = (sArr(i).Y - bHole.Y) / d
            sArr(i).X = sArr(i).X - dx
            sArr(i).vx = sArr(i).vx - dx
            sArr(i).vy = sArr(i).vy - dy
        Next i
        
        
        bHole.Y = bHole.Y + bHole.vy
        If bHole.Y > buff.Height Then hbH = False
    End If
    
    If sNum < 100 Then
        genStar = Int(Rnd * 10)
    
        If genStar = 5 Then
            sNum = sNum + 1
            sArr(sNum).vx = 0
            sArr(sNum).X = Int(Rnd * 640)
            sArr(sNum).Y = 0
            sArr(sNum).vy = Int(Rnd * 5 + 1)
        End If
        
        
    End If
    
    For i = 1 To sNum
        SetPixel buff.hdc, sArr(i).X, sArr(i).Y, vbWhite
        If sArr(i).vy > 3 Then
            SetPixel buff.hdc, sArr(i).X - 1, sArr(i).Y, vbWhite
            SetPixel buff.hdc, sArr(i).X + 1, sArr(i).Y, vbWhite
            SetPixel buff.hdc, sArr(i).X, sArr(i).Y - 1, vbWhite
            SetPixel buff.hdc, sArr(i).X, sArr(i).Y + 1, vbWhite
        End If
        If sArr(i).vy = 3 Then
            SetPixel buff.hdc, sArr(i).X - 1, sArr(i).Y, vbWhite
            SetPixel buff.hdc, sArr(i).X, sArr(i).Y - 1, vbWhite
        End If
        If hbH = False Then sArr(i).X = sArr(i).X + sArr(i).vx
        sArr(i).Y = sArr(i).Y + sArr(i).vy
    Next i
    For i = 1 To sNum
        If sArr(i).Y > buff.Height Then
            For j = i To sNum - 1
                sArr(j).vy = sArr(j + 1).vy
                sArr(j).vx = sArr(j + 1).vx
                sArr(j).X = sArr(j + 1).X
                sArr(j).Y = sArr(j + 1).Y
            Next j
            sNum = sNum - 1
        End If
    Next i
    If gGame Then
        If aShoot > 0 Then aShoot = aShoot - 1
        If atShoot > 0 Then atShoot = atShoot - 1
        
        For i = 1 To nShots
            myshots(i).X = myshots(i).X + myshots(i).vx
            myshots(i).Y = myshots(i).Y + myshots(i).vy
            If myshots(i).X > 800 Or myshots(i).X < -140 Or myshots(i).Y < -100 Then
                For j = i To nShots - 1
                    myshots(j).vx = myshots(j + 1).vx
                    myshots(j).vy = myshots(j + 1).vy
                    myshots(j).X = myshots(j + 1).X
                    myshots(j).Y = myshots(j + 1).Y
                Next j
                nShots = nShots - 1
            End If
        Next i
        
        
        If atShoot = 0 And nFight > 0 Then
            tsc = 0
            If ntShots < 400 Then
                atShoot = 5
                aShip.X = myShip.X: aShip.Y = myShip.Y: aShip.vx = myShip.vx: aShip.vy = myShip.vy
                For i = 1 To nFight
                    aShot.X = tShips(i).X + 1: aShot.Y = tShips(i).Y + 21: aShot.vx = tShips(i).vx / 2: aShot.vy = 8
                    For j = 1 To 10
                        If aShot.X - aShip.X < 26 And aShot.X - aShip.X > 0 And aShot.Y - aShip.Y < 21 And aShot.Y - aShip.Y > 0 Then
                            ntShots = ntShots + 1
                            tshots(ntShots).X = tShips(i).X + 1: aShot.Y = tShips(i).Y + 21: aShot.vx = tShips(i).vx: aShot.vy = 8
                            ntShots = ntShots + 1
                            tshots(ntShots).X = tShips(i).X + 20: aShot.Y = tShips(i).Y + 21: aShot.vx = tShips(i).vx: aShot.vy = 8
                            j = 10
                            i = nFight
                            tsc = 1
                        End If
                        aShot.X = aShot.X + aShot.vx
                        aShot.Y = aShot.Y + aShot.vy
                        aShip.X = aShip.X + aShip.vx
                        aShip.Y = aShip.Y + aShip.vy
                    Next j
                    aShot.X = tShips(i).X + 20: aShot.Y = tShips(i).Y + 21: aShot.vx = tShips(i).vx / 2: aShot.vy = 8
                    For j = 1 To 10
                        If aShot.X - aShip.X < 26 And aShot.X - aShip.X > 0 And aShot.Y - aShip.Y < 21 And aShot.Y - aShip.Y > 0 Then
                            ntShots = ntShots + 1
                            tshots(ntShots).X = tShips(i).X + 20: aShot.Y = tShips(i).Y + 21: aShot.vx = tShips(i).vx: aShot.vy = 8
                            ntShots = ntShots + 1
                            tshots(ntShots).X = tShips(i).X + 1: aShot.Y = tShips(i).Y + 21: aShot.vx = tShips(i).vx: aShot.vy = 8
                            j = 10
                            i = nFight
                            tsc = True
                       End If
                       aShot.X = aShot.X + aShot.vx
                        aShot.Y = aShot.Y + aShot.vy
                        aShip.X = aShip.X + aShip.vx
                        aShip.Y = aShip.Y + aShip.vy
                    Next j
                Next i
                
                If tsc = 0 Then
                    aShot.X = tShips(1).X: aShot.Y = tShips(1).Y: aShot.vx = tShips(1).vx / 2: aShot.vy = 8
                    For i = 2 To nFight
                        If Abs(tShips(i).X - myShip.X) < Abs(aShot.X - myShip.X) Then
                            aShot.X = tShips(i).X: aShot.Y = tShips(i).Y: aShot.vx = tShips(1).vx / 2: aShot.vy = 8
                        End If
                    Next i
                    ntShots = ntShots + 1
                    tshots(ntShots).X = aShot.X + 1: tshots(ntShots).Y = aShot.Y + 21: tshots(ntShots).vx = aShot.vx: tshots(ntShots).vy = aShot.vy
                End If
            End If
        End If
        
        
        For i = 1 To ntShots
            If tshots(i).X > 640 Or tshots(i).X < 0 Or tshots(i).Y > 480 Then
                For j = i To ntShots - 1
                    tshots(j).X = tshots(j + 1).X
                    tshots(j).Y = tshots(j + 1).Y
                    tshots(j).vx = tshots(j + 1).vx
                    tshots(j).vy = tshots(j + 1).vy
                Next j
                ntShots = ntShots - 1
            End If
        Next i
        If GetAsyncKeyState(kCtrl) And aShoot = 0 And nFight > 0 Then
            If nShots < 198 Then
                aShoot = 7
                nShots = nShots + 1
                myshots(nShots).X = myShip.X + 1
                myshots(nShots).Y = myShip.Y - 29
                myshots(nShots).vy = -8
                myshots(nShots).vx = myShip.vx / 2
                nShots = nShots + 1
                myshots(nShots).X = myShip.X + 20
                myshots(nShots).Y = myShip.Y - 29
                myshots(nShots).vy = -8
                myshots(nShots).vx = myShip.vx / 2
            End If
        End If
        If GetAsyncKeyState(kLeft) Then
            If myShip.vx > -10 Then myShip.vx = myShip.vx - 2
        End If
        If GetAsyncKeyState(kRight) Then
            If myShip.vx < 10 Then myShip.vx = myShip.vx + 2
        End If
        If GetAsyncKeyState(kUp) Then
            If myShip.vy < 6 Then myShip.vy = myShip.vy - 2
        End If
        If GetAsyncKeyState(kDown) Then
            If myShip.vy > -6 Then myShip.vy = myShip.vy + 2
        End If
        
        myShip.X = myShip.X + myShip.vx
        myShip.Y = myShip.Y + myShip.vy
        If myShip.vx <> 0 Then myShip.vx = myShip.vx - (myShip.vx / Abs(myShip.vx))
        If myShip.vy <> 0 Then myShip.vy = myShip.vy - (myShip.vy / Abs(myShip.vy))
    
        If myShip.Y > 459 Then
            myShip.vy = 0
            myShip.Y = 459
        End If
        If myShip.Y < 300 Then
            myShip.vy = 0
            myShip.Y = 300
        End If
        
    
        If myShip.X < 0 Then
            myShip.vx = 0
            myShip.X = 0
        End If
        If myShip.X > 610 Then
            myShip.vx = 0
            myShip.X = 610
        End If
    
    
        For i = 1 To nFight
            tsc = 0
            tX = 0
            tShips(i).X = tShips(i).X + tShips(i).vx
            tShips(i).Y = tShips(i).Y + tShips(i).vy
            For j = 1 To nShots
                aShot.X = myshots(j).X
                aShot.Y = myshots(j).Y
                aShot.vx = myshots(j).vx
                aShot.vy = myshots(j).vy
                aShip.X = tShips(i).X
                aShip.Y = tShips(i).Y
                aShip.vx = tShips(i).vx
                aShip.vy = tShips(i).vx
                For k = 1 To 10
                    If aShot.X - aShip.X < 26 And aShot.X - aShip.X > 0 And aShot.Y - aShip.Y < 21 And aShot.Y - aShip.Y > 0 Then
                        If aShot.vx > 0 Then
                            tX = tX - 4 / k
                        Else
                            tX = tX + 4 / k
                        End If
                        k = 5
                        tsc = 1
                    End If
                    aShot.X = aShot.X + aShot.vx
                    aShot.Y = aShot.Y + aShot.vy
                    aShip.X = aShip.X + aShip.vx
                    aShip.Y = aShip.Y + aShip.vy
                Next k
            Next j
            If tShips(i).d <> 0 Then tsc = 2
            If tsc = 0 Then
                'If tShips(i).X <> myShip.X And tShips(i).vx > -10 And tShips(i).vx < 10 Then
                '    tShips(i).vx = tShips(i).vx - (myShip.X - tShips(i).X) / (Abs(myShip.X - tShips(i).X))
                'End If
                Randomize
                tShips(i).vx = tShips(i).vx - 2 + Int(Rnd * 5)
                tShips(i).vy = tShips(i).vy - 1 + Int(Rnd * 3)
            Else
                If tsc = 1 Then
                    tShips(i).vx = tShips(i).vx + tX
                    tShips(i).vy = tShips(i).vy - 1
                Else
                    If tShips(i).d = 1 Then
                    
                    End If
                End If
            End If
            If tShips(i).vx > 10 Then tShips(i).vx = 10
            If tShips(i).vx < -10 Then tShips(i).vx = -10
            If tShips(i).vy > 5 Then tShips(i).vy = 5
            If tShips(i).vy < -5 Then tShips(i).vy = -5
            If tShips(i).X > 620 Then
                tShips(i).X = tShips(i).X - 20
                tShips(i).vx = -tShips(i).vx
            End If
            If tShips(i).X < 0 Then
                tShips(i).X = 0
                tShips(i).vx = -tShips(i).vx
            End If
            If tShips(i).Y < 0 Then
                tShips(i).Y = 0
                tShips(i).vy = 1
            End If
            If tShips(i).Y > 300 Then
                tShips(i).Y = 300
                tShips(i).vy = -tShips(i).vy
            End If
            For j = 1 To nShots
                If myshots(j).X - tShips(i).X < 26 And myshots(j).X - tShips(i).X > 0 And myshots(j).Y - tShips(i).Y < 21 And myshots(j).Y - tShips(i).Y > 0 Then
                If myshots(j).X - tShips(i).X > 8 And myshots(j).X - tShips(i).X < 17 Then
                    For k = j To nShots - 1
                        myshots(k).vx = myshots(k + 1).vx
                        myshots(k).vy = myshots(k + 1).vy
                        myshots(k).X = myshots(k + 1).X
                        myshots(k).Y = myshots(k + 1).Y
                    Next k
                    nShots = nShots - 1
                    For k = i To nFight - 1
                        tShips(k).vx = tShips(k + 1).vx
                        tShips(k).vy = tShips(k + 1).vy
                        tShips(k).X = tShips(k + 1).X
                        tShips(k).Y = tShips(k + 1).Y
                    Next k
                'buff.BackColor = vbWhite
                nFight = nFight - 1
                End If
                
                If myshots(j).X - tShips(i).X < 9 Then
                
                    If tShips(i).t = 0 Then
                        tShips(i).t = 1
                    End If
                
                    If tShips(i).t = 2 Then
                        For k = j To nShots - 1
                            myshots(k).vx = myshots(k + 1).vx
                            myshots(k).vy = myshots(k + 1).vy
                            myshots(k).X = myshots(k + 1).X
                            myshots(k).Y = myshots(k + 1).Y
                        Next k
                        nShots = nShots - 1
                        For k = i To nFight
                            tShips(k).vx = tShips(k + 1).vx
                            tShips(k).vy = tShips(k + 1).vy
                            tShips(k).X = tShips(k + 1).X
                            tShips(k).Y = tShips(k + 1).Y
                        Next k
                        buff.BackColor = vbWhite
                        nFight = nFight - 1
                    
                    End If
                End If
                
                If myshots(j).X - tShips(i).X > 16 Then
                
                    If tShips(i).t = 0 Then
                        tShips(i).t = 2
                    End If
                
                    If tShips(i).t = 1 Then
                        For k = j To nShots - 1
                            myshots(k).vx = myshots(k + 1).vx
                            myshots(k).vy = myshots(k + 1).vy
                            myshots(k).X = myshots(k + 1).X
                            myshots(k).Y = myshots(k + 1).Y
                        Next k
                        nShots = nShots - 1
                        For k = i To nFight
                            tShips(k).vx = tShips(k + 1).vx
                            tShips(k).vy = tShips(k + 1).vy
                            tShips(k).X = tShips(k + 1).X
                            tShips(k).Y = tShips(k + 1).Y
                        Next k
                        buff.BackColor = vbWhite
                        nFight = nFight - 1
                    
                    End If
                End If
                
                End If
            Next j
        Next i
        
        If nFight > 3 Then
            Randomize
            d = Int(Rnd * 1500)
            If d = 1 Then
                tShips(1).d = 1
                tShips(2).d = 2
                tShips(3).d = 3
            End If
        End If
        
        For i = 0 To 25
            For j = 0 To 20
                If mFire(i, j) Then SetPixel ship.hdc, i, j, GetPixel(pal.hdc, Int(Rnd * 5), 0)
            Next j
        Next i
        For i = 0 To 25
            For j = 0 To 20
            
            Next j
        Next i
        
        BitBlt buff.hdc, myShip.X, myShip.Y, ship.Width, ship.Height, ship.hdc, 0, 0, SRCPAINT
        
        For i = 1 To ntShots
            tshots(i).X = tshots(i).X + tshots(i).vx
            tshots(i).Y = tshots(i).Y + tshots(i).vy
            BitBlt buff.hdc, tshots(i).X, tshots(i).Y, 5, 8, eShot.hdc, Int(Rnd * 3) * 5, 0, SRCPAINT
            If tshots(i).X - myShip.X < 26 And tshots(i).X - myShip.X > -5 And tshots(i).Y - myShip.Y < 21 And tshots(i).Y - myShip.Y > -7 Then
                For k = 0 To 5
                    For l = 2 To 7
                        If GetPixel(ship.hdc, k + tshots(i).X - myShip.X, l + tshots(i).Y - myShip.Y) <> vbBlack And GetPixel(eShot.hdc, k, l) <> vbBlack Then
                            On Error Resume Next
                            mFire(k + tshots(i).X - myShip.X, l + tshots(i).Y - myShip.Y) = True
                            'myShip.X = 307
                            'myShip.Y = 480 - ship.Height
                            'myShip.vx = 0
                            'myShip.vy = 0
                            'ntShots = 0
                            'nShots = 0
                            'nFight = 0
                            'nLevel = 0
                            'MsgBox "You Lose!"
                            'GoTo loopend
                        End If
                    Next l
                Next k
            End If
loopend:
            'If tshots(i).X - myShip.X < 26 And tshots(i).X - myShip.X > 0 And tshots(i).Y - myShip.Y < 21 And tshots(i).Y - myShip.Y > 0 Then
            '    myShip.X = 307
            '    myShip.Y = 480 - ship.Height
            '    myShip.vx = 0
            '    myShip.vy = 0
            '    ntShots = 0
            '    nShots = 0
            '    nFight = 0
            '    nLevel = 0
            '    MsgBox "You Lose!"
                'gGame = False
            'End If
        Next i
        
        For i = 1 To nShots
            BitBlt buff.hdc, myshots(i).X, myshots(i).Y, 5, 8, shot.hdc, Int(Rnd * 3) * 5, 0, SRCPAINT
        Next i
        
        For i = 1 To nFight
            BitBlt buff.hdc, tShips(i).X, tShips(i).Y, 26, 21, eShip(tShips(i).t).hdc, 0, 0, SRCPAINT
        Next i
        
        Randomize
        For i = 1 To 4
            For j = 40 To 2 Step -1
                Fire(i, j) = Fire(i, j - 1) + Int(Rnd * 2)
            Next j
        Next i
        For i = 1 To 4
            Fire(i, 1) = Int(Rnd * 3)
        Next i
        
        For i = 1 To 4
            For j = 1 To 10
                If (j - 5) < -myShip.vy Then
                    If Fire(i, j) > 5 Then
                          SetPixel buff.hdc, myShip.X + i + 10, myShip.Y + j + 20, vbBlack
                    Else
                        SetPixel buff.hdc, myShip.X + i + 10, myShip.Y + j + 20, GetPixel(pal.hdc, Fire(i, j), 0)
                    End If
                End If
            Next j
        Next i
        
    End If
    j = 1 + nFight + nShots + ntShots + hbH
    For i = 1 To 10
        If myFight(i).isInit Then
            j = j + 2 + myFight(i).nShots
        End If
    Next i
    If Second(Time) <> sec Then
        buff.Print "FPS: " & frame
        buff.Print "SPRITES: " & (j)
        lframe = frame
        frame = 0
        sec = Second(Time)
    Else
        frame = frame + 1
        buff.Print "FPS: " & lframe
        buff.Print "SPRITES: " & (j)
    End If
        BitBlt Me.hdc, (Me.Width / Screen.TwipsPerPixelX - buff.Width) / 2, (Me.Height / Screen.TwipsPerPixelY - buff.Height) / 2, buff.Width, buff.Height, buff.hdc, 0, 0, SRCCOPY
    Randomize
Else
sCounter = sCounter + 1
End If
Loop

End Sub

Function fAdd(num As Integer) As Integer
Dim i As Integer, j As Integer: j = 0
If num > 0 Then
    For i = num To 1 Step -1
        j = j + i
    Next i
Else
    For i = num To -1 Step 1
        j = j + i
    Next i
End If
fAdd = j
End Function


