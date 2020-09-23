VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   4575
      Left            =   0
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b(20) As ObjectB    'this holds the ball information for 31 balls
Dim Running As Boolean
Dim V As Single
Sub Run()
    
    While Running
        P.Cls
        V = 0
        For i = 0 To UBound(b)
            'comment out the next line to remove the effect of friction
            HandleFriction b(i)
            
            'deal with sides of the picture box
            If b(i).X - b(i).Radius < 0 Then
                b(i).Vx = -b(i).Vx
                b(i).X = b(i).Radius
            End If
            If b(i).X + b(i).Radius > P.ScaleWidth Then
                b(i).Vx = -b(i).Vx
                b(i).X = P.ScaleWidth - b(i).Radius
            End If
            If b(i).Y - b(i).Radius < 0 Then
                b(i).Vy = -b(i).Vy
                b(i).Y = b(i).Radius
            End If
            If b(i).Y + b(i).Radius > P.ScaleHeight Then
                b(i).Vy = -b(i).Vy
                b(i).Y = P.ScaleHeight - b(i).Radius
            End If
            '____________________________________
             
            'check for collision between current ball and other balls
            For j = 0 To UBound(b)
                If i <> j Then ChangeVelocities b(j), b(i)
            Next j
            'update the position of the current ball
            b(i).X = b(i).Vx + b(i).X
            b(i).Y = b(i).Vy + b(i).Y
            'draw the ball
            P.Circle (b(i).X, b(i).Y), b(i).Radius
            'V = V + (b(i).Vx ^ 2 + b(i).Vy ^ 2) * 0.5 * b(i).Mass
        Next i
        

        'delay loop
        For i = 1 To 10000
        Next i
        DoEvents
    Wend
End Sub

Private Sub Command1_Click()
    For i = 0 To UBound(b)
        If i <> 0 Then b(i).Radius = Int(Rnd * 10) + 5
        b(i).Vx = (Rnd * 4) - 2
        b(i).Vy = (Rnd * 4) - 2
        b(i).X = (Rnd * (P.ScaleWidth - 2 * b(i).Radius)) + b(i).Radius
        b(i).Y = (Rnd * (P.ScaleHeight - 2 * b(i).Radius)) + b(i).Radius
        b(i).Mass = b(i).Radius
    Next i
    
    b(0).Radius = 50
    b(0).Mass = b(0).Radius
    Running = True
    Run
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Running = False
End Sub

