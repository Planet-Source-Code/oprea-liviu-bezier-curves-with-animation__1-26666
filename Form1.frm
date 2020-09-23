VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BEZIER CURVE by Oprea Liviu"
   ClientHeight    =   4410
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuDraw 
         Caption         =   "&Draw Bezier "
      End
      Begin VB.Menu mnuAnimate 
         Caption         =   "&Animate"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop Animation"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "&Choose Color"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private m_DraggingHandle As Integer

' The data points.
Private m_NumPoints As Single
Private m_PointX() As Single
Private m_PointY() As Single

Dim lppt() As POINTAPI
Dim cPoints As Long
Dim curba As Boolean
Dim m As Boolean
Dim DeltaX As Integer
Dim DeltaY As Integer
Dim b() As POINTAPI
Private Const HANDLE_WIDTH = 8
Private Const HANDLE_HALF_WIDTH = HANDLE_WIDTH / 2


' Create some initial point data.
Private Sub Form_Load()
  Dim i As Integer
    Me.ScaleMode = vbPixels
    Me.BackColor = vbWhite
    m_NumPoints = 7
    ReDim m_PointX(1 To m_NumPoints)
    ReDim m_PointY(1 To m_NumPoints)
    curba = False
    ' Set initial points.
    m_PointX(1) = 320: m_PointY(1) = 144
    m_PointX(2) = 545: m_PointY(2) = 510
    m_PointX(3) = 240: m_PointY(3) = 95
    m_PointX(4) = 80: m_PointY(4) = 380
    m_PointX(5) = 460: m_PointY(5) = 414
    m_PointX(6) = 90: m_PointY(6) = 250
    m_PointX(6) = 300: m_PointY(7) = 300
    ' Draw.
    Refresh
    Timer1.Interval = 1
    DeltaX = 5
    DeltaY = 5
    
ReDim b(m_NumPoints) As POINTAPI
For i = 1 To m_NumPoints
  If (i Mod 2 = 0) Then
    b(i).x = 1
    b(i).y = -1
  Else
  b(i).x = -1
  b(i).y = 1
 End If
Next i
 Timer1.Enabled = False
 mnuStop.Enabled = False
 c = 12
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
Dim dx As Single
Dim dy As Single

    For i = 1 To m_NumPoints
        If Abs(m_PointX(i) - x) < HANDLE_HALF_WIDTH And _
           Abs(m_PointY(i) - y) < HANDLE_HALF_WIDTH _
        Then
            ' We are over this grab handle.
            ' Start dragging.
            m_DraggingHandle = i
            Exit For
        End If
    Next i
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Do nothing if we are not dragging.
    Dim i As Integer
    If m_DraggingHandle = 0 Then Exit Sub
    
    ' Move the handle.
    m_PointX(m_DraggingHandle) = x
    m_PointY(m_DraggingHandle) = y
    
    ' Redraw.
  Refresh
End Sub


' Stop dragging
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_DraggingHandle = 0
End Sub


Private Sub Form_Paint()
Dim i As Integer

  
    If m_NumPoints < 1 Then Exit Sub

    ' Porneste de la ultimul punct
    CurrentX = m_PointX(m_NumPoints)
    CurrentY = m_PointY(m_NumPoints)

    ' Connect the points.
   

    
    FillColor = vbWhite
    FillStyle = vbFSSolid
    
    If curba Then
     draw (4)
    Else
      For i = 1 To m_NumPoints
        Line -(m_PointX(i), m_PointY(i))
    Next i
    
    End If
    If Not m Then
    For i = 1 To m_NumPoints
        Line (m_PointX(i) - HANDLE_HALF_WIDTH, m_PointY(i) - HANDLE_HALF_WIDTH)-Step(HANDLE_WIDTH, HANDLE_WIDTH), , B
    Next i
    End If
End Sub



Private Sub draw(culoare As Integer)
Dim hDC As Long
Dim lppt() As POINTAPI
Dim cPoints As Long
Dim i As Integer
Dim j As Integer
ReDim lppt(3 * m_NumPoints + 1)

j = 2
Me.ForeColor = QBColor(culoare)
hDC = Me.hDC
lppt(1).x = (m_PointX(1) + m_PointX(2)) / 2
lppt(1).y = (m_PointY(1) + m_PointY(2)) / 2

For i = 1 To m_NumPoints
   lppt(j).x = m_PointX(i Mod m_NumPoints + 1)
   lppt(j).y = m_PointY(i Mod m_NumPoints + 1)
   j = j + 1
   
    lppt(j).x = m_PointX(i Mod m_NumPoints + 1)
    lppt(j).y = m_PointY(i Mod m_NumPoints + 1)
    j = j + 1
 
    lppt(j).x = (m_PointX(i Mod m_NumPoints + 1) + _
        m_PointX((i + 1) Mod m_NumPoints + 1)) / 2
    lppt(j).y = (m_PointY(i Mod m_NumPoints + 1) + _
        m_PointY((i + 1) Mod m_NumPoints + 1)) / 2
    j = j + 1
  Next i
 Me.DrawWidth = 2
 Call PolyBezier(hDC, lppt(1), 3 * m_NumPoints + 1)
curba = True
Me.DrawWidth = 1
End Sub

Private Sub mnuAnimate_Click()
m = True
Timer1.Enabled = True
mnuStop.Enabled = True
mnuAnimate.Enabled = False
End Sub

Private Sub mnuColor_Click()
 Form2.Show
 
End Sub

Private Sub mnuDraw_Click()
draw (12)
curba = True

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuStop_Click()
Timer1.Enabled = False
mnuStop.Enabled = False
mnuAnimate.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim i As Integer

Me.BackColor = vbBlack
draw (0)
For i = 1 To m_NumPoints
 m_PointX(i) = m_PointX(i) + DeltaX * b(i).x
 m_PointY(i) = m_PointY(i) + DeltaY * b(i).y
 If m_PointX(i) < ScaleLeft Then b(i).x = 1
 If m_PointX(i) > ScaleLeft + ScaleWidth Then b(i).x = -1
 If m_PointY(i) < ScaleTop Then b(i).y = 1
 If m_PointY(i) > ScaleHeight + ScaleTop Then b(i).y = -1
Next i
 draw (c)
  'Refresh
End Sub
