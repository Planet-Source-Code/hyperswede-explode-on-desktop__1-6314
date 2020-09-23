VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawStyle       =   5  'Transparent
   DrawWidth       =   10
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
 Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long





Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        Y As Long
End Type
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4


Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


Dim Xcur As Single, Ycur As Single
Const DE As Boolean = True
Dim SizeMod As Single
Dim CurExplo As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type tpBoom
Size As Integer
Heading As Double
Speed As Integer
CurX As Single
CurY As Single
GoingDown As Boolean
Done As Boolean
MaxSize As Integer
End Type
Private Type tpDebree
tmp As String
End Type
Private Type tpSpark
tmp As String
End Type
Private Type typeBoomInfo
Ammount As Integer
MaxOutSpeed As Integer
MinOutSpeed As Integer
SizeAccelerationModifier As Single
SizeDecelerationModifier As Single
StartSize As Integer
MaxSize As Integer
MinMaxSize As Integer
color As Integer
MinStartSize As Integer
End Type
Private Type typeDebreeInfo
Ammount As Integer
Weight As Integer
MinSize As Integer
MaxSize As Integer
End Type
Private Type typeSparksInfo
Ammount As Integer
StartR As Integer
StartG As Integer
StartB As Integer
CoolSpeed As Integer
End Type
Private Type tpExplosion
Boom(1 To 1000) As tpBoom
Xpos As Single
Ypos As Single
End Type
Dim BoomStuff As typeBoomInfo, DebreeStuff As typeDebreeInfo, SparkStuff As typeSparksInfo
Private Type Tx
BoomColor(1 To 1000) As Integer
End Type
Dim BC As Tx
Dim BCG As Tx
Dim BCB As Tx
Const MaxExplosions As Integer = 1
Dim Explosion(1 To MaxExplosions) As tpExplosion
Private Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
    ' This function uses an argument to determine whet    '     her
    '     ' to make the specified form always on top or not
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop Then
        OnTop = SetWindowPos(FrmID.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(FrmID.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Private Function BetweenRND(lowerbound As Integer, upperbound As Integer) As Integer
BetweenRND = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Private Sub Explode(Xp As Single, Yp As Single, BoomInfo As typeBoomInfo, DebreeInfo As typeDebreeInfo, SparksInfo As typeSparksInfo)
Dim ntFirst As Boolean
CurExplo = CurExplo + 1
If Int(CurExplo) > Int(MaxExplosions) Then
CurExplo = CurExplo - 1
Exit Sub
End If
FillColor = BoomInfo.color
Dim StopExplode As Boolean, BigX As Single, BigY As Single, SmallX As Single, SmallY As Single
Explosion(CurExplo).Xpos = Xp
Explosion(CurExplo).Ypos = Yp
For i = 1 To BoomInfo.Ammount
Explosion(CurExplo).Boom(i).Size = BetweenRND(BoomInfo.MinStartSize, BoomInfo.StartSize)
Explosion(CurExplo).Boom(i).Heading = Rnd * 6.28
Explosion(CurExplo).Boom(i).MaxSize = BetweenRND(BoomInfo.MinMaxSize, BoomInfo.MaxSize)
'Explosion(curexplo).Boom(i).Speed = (BoomInfo.MaxSize - Explosion(curexplo).Boom(i).MaxSize) * BetweenRND(BoomInfo.MinOutSpeed, BoomInfo.MaxOutSpeed) / BoomInfo.MaxOutSpeed
Explosion(CurExplo).Boom(i).Speed = BetweenRND(BoomInfo.MinOutSpeed / Explosion(CurExplo).Boom(i).MaxSize * BoomInfo.MaxSize, BoomInfo.MaxOutSpeed / Explosion(CurExplo).Boom(i).MaxSize * BoomInfo.MaxSize) * (Rnd * 0.2 + 1)
Explosion(CurExplo).Boom(i).CurX = Explosion(CurExplo).Xpos
Explosion(CurExplo).Boom(i).CurY = Explosion(CurExplo).Ypos
Explosion(CurExplo).Boom(i).GoingDown = False
Explosion(CurExplo).Boom(i).Done = False
BC.BoomColor(i) = BoomInfo.color * i / BoomInfo.Ammount * 2.5
BCB.BoomColor(i) = 255 * (BoomInfo.MaxOutSpeed - Explosion(CurExplo).Boom(i).Speed) / BoomInfo.MaxOutSpeed * i / BoomInfo.Ammount * 1.5
If BCB.BoomColor(i) < 0 Then BCB.BoomColor(i) = 0
If BC.BoomColor(i) > 255 Then BC.BoomColor(i) = 255
Dim tc As Integer
tc = Int(BC.BoomColor(i) * BetweenRND(75, 80) / 100)
BCG.BoomColor(i) = tc
Next i
Dim CtC As Integer
CtC = CurExplo
Dim XBigX As Single, XBigY As Single, XSmallX As Single, XSmallY As Single
XSmallX = ScaleWidth
XSmallY = scaleheigth
Do

If DE Then DoEvents
Cls
'If DE Then Line (SmallX - BoomStuff.MaxSize, SmallY - BoomStuff.MaxSize)-(BigX + BoomStuff.MaxSize, BoomStuff.MaxSize + BigY), vbBlack, BF
Dim DoneAll As Boolean
DoneAll = True
SmallX = ScaleWidth
SmallY = scaleheigth
For i = 1 To BoomInfo.Ammount

If Not Explosion(CtC).Boom(i).GoingDown Then
Explosion(CtC).Boom(i).Size = Explosion(CtC).Boom(i).Size + (BoomInfo.MaxSize - Explosion(CtC).Boom(i).Size) / BoomInfo.SizeAccelerationModifier + 15
If Explosion(CtC).Boom(i).Size >= Explosion(CtC).Boom(i).MaxSize Then Explosion(CtC).Boom(i).GoingDown = True
Else
Explosion(CtC).Boom(i).Size = Explosion(CtC).Boom(i).Size - Explosion(CtC).Boom(i).Size / BoomInfo.SizeDecelerationModifier - 1
If Explosion(CtC).Boom(i).Size > 0 Then Explosion(CtC).Boom(i).Done = False Else Explosion(CtC).Boom(i).Done = True
End If
Explosion(CtC).Boom(i).Speed = Explosion(CtC).Boom(i).Speed - 1
If Explosion(CtC).Boom(i).Speed < 0 Then Explosion(CtC).Boom(i).Speed = 0
Explosion(CtC).Boom(i).CurX = Explosion(CtC).Boom(i).CurX + Cos(Explosion(CtC).Boom(i).Heading) * Explosion(CtC).Boom(i).Speed
Explosion(CtC).Boom(i).CurY = Explosion(CtC).Boom(i).CurY + Sin(Explosion(CtC).Boom(i).Heading) * Explosion(CtC).Boom(i).Speed
ptx = Explosion(CtC).Boom(i).CurX
pty = Explosion(CtC).Boom(i).CurY
If ptx < 0 Or pty < 0 Or ptx > ScaleWidth Or pty > ScaleHeight Then
Explosion(CtC).Boom(i).Heading = Explosion(CtC).Boom(i).Heading + 3.14
Explosion(CtC).Boom(i).Speed = Explosion(CtC).Boom(i).Speed * 0
Explosion(CtC).Boom(i).CurX = ptx: Explosion(CtC).Boom(i).CurY = pty
End If
If Explosion(CtC).Boom(i).Done = False Then
FillColor = RGB(BC.BoomColor(i), BCG.BoomColor(i), BCB.BoomColor(i))
DrawMode = 15
Me.Circle (Explosion(CtC).Boom(i).CurX, Explosion(CtC).Boom(i).CurY), Explosion(CtC).Boom(i).Size, FillColor
'Me.PSet (Explosion(CtC).Boom(i).CurX, Explosion(CtC).Boom(i).CurY), FillColor
If Not BCG.BoomColor(i) <= 135 Then
BCG.BoomColor(i) = BCG.BoomColor(i) - Explosion(CtC).Boom(i).Speed / 2 - 4
If Not BCB.BoomColor(i) <= 30 Then BCB.BoomColor(i) = BCB.BoomColor(i) - 0
BC.BoomColor(i) = BC.BoomColor(i) - 1
Else
BCG.BoomColor(i) = BCG.BoomColor(i) - 20
BC.BoomColor(i) = BC.BoomColor(i) - 20
BCB.BoomColor(i) = BCB.BoomColor(i) - 30
If BCB.BoomColor(i) < 0 Then BCB.BoomColor(i) = 0
If BC.BoomColor(i) < 0 Then BC.BoomColor(i) = 0
If BCG.BoomColor(i) < 0 Then BCG.BoomColor(i) = 0
End If
End If
If BC.BoomColor(i) = 0 And BCG.BoomColor(i) = 0 And BCB.BoomColor(i) = 0 Then Explosion(CtC).Boom(i).Done = True

If Explosion(CtC).Boom(i).Done = False Then DoneAll = False
If Explosion(CtC).Boom(i).CurX > BigX Then BigX = Explosion(CtC).Boom(i).CurX
If Explosion(CtC).Boom(i).CurY > BigY Then BigY = Explosion(CtC).Boom(i).CurY
If Explosion(CtC).Boom(i).CurX < SmallX Then SmallX = Explosion(CtC).Boom(i).CurX
If Explosion(CtC).Boom(i).CurY < SmallY Then SmallY = Explosion(CtC).Boom(i).CurY
If BigX > XBigX Then XBigX = BigX: If BigY > XBigY Then XBigY = BigY: If SmallX < XSmallX Then XSmallX = SmallX: If SmallY < XSmallY Then XSmallY = SmallY
Next i
If DoneAll = True Then StopExplode = True
If DE Then DoEvents
Loop Until StopExplode
If DE Then Line (SmallX - BoomStuff.MaxSize, SmallY - BoomStuff.MaxSize)-(BigX + BoomStuff.MaxSize, BoomStuff.MaxSize + BigY), vbBlack, BF

CurExplo = CurExplo - 1
Cls
End Sub

Private Sub Form_Activate()
AlwaysOnTop Me, True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
ElseIf KeyCode = vbKeyReturn Then
frmBombSel.Show
AlwaysOnTop frmBombSel, True
Do: DoEvents: Loop Until frmBombSel.Visible = False
SizeMod = CSng(frmBombSel.Tag)
Else
Cls
End If
End Sub

Private Sub Form_Load()
SizeMod = 2.5
Move Screen.Width, Screen.Height, 0, 0
Randomize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Xcur = x
Ycur = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer_Timer()
If GetAsyncKeyState(VK_MBUTTON) < 0 Then
If GetAsyncKeyState(VK_LBUTTON) < 0 Then
Dim Cpos As POINTAPI
GetCursorPos Cpos
Move 0, 0, Screen.Width, Screen.Height
Hide
        DoEvents
        hw = GetDesktopWindow()
        ha = GetDC(hw)
        thx = Left / Screen.TwipsPerPixelX
        thy = Top / Screen.TwipsPerPixelY
        wdth = ScaleWidth
        higt = ScaleHeight
        Show
        tmp = BitBlt(hdc, 0, 0, wdth, higt, ha, thx, thy, SRCCOPY)
        Picture = Image
        Call ReleaseDC(hw, ha)
 
       
BoomStuff.Ammount = 40 '* SizeMod
BoomStuff.color = 255
BoomStuff.MaxOutSpeed = 9
BoomStuff.MinOutSpeed = 6
BoomStuff.MaxSize = 15 * SizeMod
BoomStuff.SizeAccelerationModifier = 40 '* SizeMod
BoomStuff.SizeDecelerationModifier = 20 '* SizeMod
BoomStuff.StartSize = 2 '* SizeMod
BoomStuff.MinStartSize = 1 '* SizeMod
BoomStuff.MinMaxSize = 5 * SizeMod
Explode CSng(Cpos.x), CSng(Cpos.Y), BoomStuff, DebreeStuff, SparkStuff

Move Screen.Width, Screen.Height, 0, 0
ElseIf GetAsyncKeyState(VK_RBUTTON) < 0 Then
Form_KeyDown vbKeyReturn, 0
End If
End If
End Sub
