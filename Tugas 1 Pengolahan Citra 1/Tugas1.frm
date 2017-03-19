VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tugas 1"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   12480
      Max             =   20
      Min             =   1
      TabIndex        =   12
      Top             =   6360
      Value           =   10
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   12480
      Max             =   127
      Min             =   -127
      TabIndex        =   10
      Top             =   5160
      Width           =   2175
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Left            =   3120
      ScaleHeight     =   2955
      ScaleWidth      =   4275
      TabIndex        =   9
      Top             =   5160
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Kontras"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Brightness"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Histogram"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   3015
      Left            =   7680
      ScaleHeight     =   2955
      ScaleWidth      =   4275
      TabIndex        =   5
      Top             =   5160
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Negative"
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Binary"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gray-Scale"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   120
      Picture         =   "Tugas1.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   7155
      TabIndex        =   1
      Top             =   120
      Width           =   7215
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   7680
      ScaleHeight     =   3675
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label6 
      Caption         =   "Adjust Kontras"
      Height          =   255
      Left            =   12480
      TabIndex        =   17
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Adjust Brightness"
      Height          =   255
      Left            =   12480
      TabIndex        =   16
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Histogram Gambar Olahan"
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Histogram Gambar Asli"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   12480
      TabIndex        =   13
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   12480
      TabIndex        =   11
      Top             =   5760
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15000
      Y1              =   4680
      Y2              =   4680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    For i = 1 To Picture1.Width Step 15
        For j = 1 To Picture1.Height Step 15
        warna = Picture1.Point(i, j)
        R = warna And RGB(255, 0, 0)
        G = Int((warna And RGB(0, 255, 0)) / 256)
        B = Int(Int((warna And RGB(0, 0, 255)) / 256) / 256)
        x = 0.42 * R + 0.32 * G + 0.28 * B
        Picture2.PSet (i, j), RGB(x, x, x)
        Next j
    Next i
    MsgBox "Selesai"
End Sub

Private Sub Command2_Click()
    Dim xr As Single
    Dim wx(500, 500) As Integer
    xr = 0: m = 0
    For i = 1 To Picture1.Width Step 15
        m = m + 1
        n = 0
        For j = 1 To Picture1.Height Step 15
            warna = Picture1.Point(i, j)
            eR = warna And RGB(255, 0, 0)
            Ge = Int((warna And RGB(0, 255, 0)) / 256)
            Be = Int(Int((warna And RGB(0, 0, 255)) / 256) / 256)
            n = n + 1
            x = (eR + Ge + Be) / 3
            xr = xr + x
            wx(m, n) = x
        Next j
    Next i
    xr = xr / (m * n)
    For i = 0 To m
        For j = 0 To n
            If wx(i, j) < xr Then x = 0 Else x = 255
            Picture2.PSet (15 * (i - 1) + 1, 15 * (j - 1) + 1), RGB(x, x, x)
        Next j
    Next i
    MsgBox "Selesai"
End Sub

Private Sub Command3_Click()
    For i = 1 To Picture1.Width Step 15
        For j = 1 To Picture1.Height Step 15
        warna = Picture1.Point(i, j)
        R = warna And RGB(255, 0, 0)
        G = Int((warna And RGB(0, 255, 0)) / 256)
        B = Int(Int((warna And RGB(0, 0, 255)) / 256) / 256)
        x = (R + G + B) / 3
        xn = 255 - x
        Picture2.PSet (i, j), RGB(xn, xn, xn)
        Next j
    Next i
    MsgBox "Selesai"
End Sub

Private Sub Command4_Click()
    Dim h1(256) As Double
    Dim h2(256) As Double
    Picture3.Cls
    Picture4.Cls
    For i = 0 To 255
        h1(i) = 0
        h2(i) = 0
    Next i
    For i = 1 To Picture2.Width Step 15
        For j = 1 To Picture2.Height Step 15
        warna = Picture2.Point(i, j)
        R = warna And RGB(255, 0, 0)
        G = Int((warna And RGB(0, 255, 0)) / 256)
        B = Int(Int((warna And RGB(0, 0, 255)) / 256) / 256)
        x = Int((R + G + B) / 3)
        h1(x) = h1(x) + 1
        warna = Picture1.Point(i, j)
        R = warna And RGB(255, 0, 0)
        G = Int((warna And RGB(0, 255, 0)) / 256)
        B = Int(Int((warna And RGB(0, 0, 255)) / 256) / 256)
        x = Int((R + G + B) / 3)
        h2(x) = h2(x) + 1
        Next j
    Next i
    ht2 = Picture3.Height
    For i = 0 To 256
        xp = 15 * (i) + 1
        Picture3.Line (xp, ht2 - h1(i))-(xp, ht2), RGB(0, 255, 0)
        Picture4.Line (xp, ht2 - h2(i))-(xp, ht2), RGB(255, 255, 0)
    Next i
    MsgBox "Selesai"
End Sub

Private Sub Command5_Click()
    nl = Val(HScroll1)
    For i = 1 To Picture1.Width Step 15
        For j = 1 To Picture1.Height Step 15
        warna = Picture1.Point(i, j)
        R = warna And RGB(255, 0, 0)
        G = Int((warna And RGB(0, 255, 0)) / 256)
        B = Int(Int((warna And RGB(0, 0, 255)) / 256) / 256)
        x = (R + G + B) / 3
        xr = x + nl
        If xr < 0 Then xr = 0
        If xr > 255 Then xr = 255
        Picture2.PSet (i, j), RGB(xr, xr, xr)
        Next j
    Next i
    MsgBox "Selesai"
End Sub

Private Sub Command6_Click()
    nl = Val(HScroll2) / 10
    For i = 1 To Picture1.Width Step 15
        For j = 1 To Picture1.Height Step 15
        warna = Picture1.Point(i, j)
        R = warna And RGB(255, 0, 0)
        G = Int((warna And RGB(0, 255, 0)) / 256)
        B = Int(Int((warna And RGB(0, 0, 255)) / 256) / 256)
        x = (R + G + B) / 3
        xr = x * nl
        If xr < 0 Then xr = 0
        If xr > 255 Then xr = 255
        Picture2.PSet (i, j), RGB(xr, xr, xr)
        Next j
    Next i
    MsgBox "Selesai"
End Sub

Private Sub HScroll1_Change()
    Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
    Label3.Caption = HScroll2.Value
End Sub
