VERSION 5.00
Begin VB.Form frmboggletest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Algorithm Optimizer"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBlock 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Shape shpA 
      BackColor       =   &H00FFFF00&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   15
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Optimizing Algorithms"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmboggletest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AA(5), Z As Integer
Dim POON0(10000), POON1(10000), POON2(10000)
Dim POON3(10000), POON4(10000), POON5(10000)
Dim DaCOUNT, fPOS, fLEN, tmpHALF
Dim PCTtotal As Long
Dim TIMEsoFAR As Integer
Private Sub BooT()
Dim ARSE
Dim BOOGie As Boolean
ARSE = FreeFile
Open App.Path & "\algorithms.dat" For Append As #ARSE
On Error Resume Next
For Z = DaCOUNT To 0 Step -1
    If POON0(Z) = POON1(Z) Then BOOGie = False
    If POON0(Z) = POON2(Z) Then BOOGie = False
    If POON0(Z) = POON3(Z) Then BOOGie = False
    If POON0(Z) = POON4(Z) Then BOOGie = False
    If POON0(Z) = POON5(Z) Then BOOGie = False
    If POON1(Z) = POON2(Z) Then BOOGie = False
    If POON1(Z) = POON3(Z) Then BOOGie = False
    If POON1(Z) = POON4(Z) Then BOOGie = False
    If POON1(Z) = POON5(Z) Then BOOGie = False
    If POON2(Z) = POON3(Z) Then BOOGie = False
    If POON2(Z) = POON4(Z) Then BOOGie = False
    If POON2(Z) = POON5(Z) Then BOOGie = False
    If POON3(Z) = POON4(Z) Then BOOGie = False
    If POON3(Z) = POON5(Z) Then BOOGie = False
    If POON4(Z) = POON5(Z) Then BOOGie = False
    'frmboggletest.Caption = Z & " to go"
    'Refresh
    If BOOGie = True Then Print #ARSE, POON0(Z) & "," & POON1(Z) & "," & POON2(Z) & "," & POON3(Z) & "," & POON4(Z) & "," & POON5(Z)
    BOOGie = True
    DoEvents
    PCTtotal = PCTtotal + tmpHALF / DaCOUNT
Next Z
Close #ARSE
DaCOUNT = 0
tmpHALF = 0
End Sub
Private Sub Form_Load()
Show
Reset
fLEN = FileLen(App.Path & "\temp.txt")
frmBoggle.Enabled = False
Open App.Path & "\temp.txt" For Input As #1
    Do Until EOF(1)
        Input #1, AA(0)
        fPOS = fPOS + Len(AA(0))
        Input #1, AA(1)
        fPOS = fPOS + Len(AA(1))
        Input #1, AA(2)
        fPOS = fPOS + Len(AA(2))
        Input #1, AA(3)
        fPOS = fPOS + Len(AA(3))
        Input #1, AA(4)
        fPOS = fPOS + Len(AA(4))
        Input #1, AA(5)
        fPOS = fPOS + Len(AA(5)) + 7
        POON0(DaCOUNT) = AA(0)
        POON1(DaCOUNT) = AA(1)
        POON2(DaCOUNT) = AA(2)
        POON3(DaCOUNT) = AA(3)
        POON4(DaCOUNT) = AA(4)
        POON5(DaCOUNT) = AA(5)
        DaCOUNT = DaCOUNT + 1
        'frmboggletest.Caption = DaCOUNT
        If DaCOUNT > 10000 Then Call BooT
        DoEvents
        tmpHALF = fPOS / 2 + tmpHALF
        PCTtotal = fPOS / 2 + PCTtotal
        fPOS = 0
        'shpA.Width = ((PCTtotal / 100) / (fLEN / 100)) * Label2.Width
        'Label3.Caption = "Block " & PCTtotal & " of " & fLEN & "  --  " & Round(((fPOS / 100) / (fLEN / 100)) * 100, 2) & "%"
    Loop
Close #1
Call BooT
Kill App.Path & "\temp.txt"
MsgBox "Algorithm file has been rebuilt.  Please re-open program."
End
End Sub

Private Sub tmrBlock_Timer()
TIMEsoFAR = TIMEsoFAR + 1
Label3.Caption = "Elapsed time: " & TIMEsoFAR & " sec   Time Left: " & Int((fLEN - PCTtotal) * (TIMEsoFAR / PCTtotal)) & " secs"
shpA.Width = ((PCTtotal / 100) / (fLEN / 100)) * Label2.Width
End Sub
