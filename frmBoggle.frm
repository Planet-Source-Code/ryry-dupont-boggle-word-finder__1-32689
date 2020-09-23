VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBoggle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boggle Word Finder"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmBoggle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Menu"
      Height          =   255
      Left            =   2760
      TabIndex        =   50
      Top             =   5400
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2040
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   5
      Left            =   3840
      TabIndex        =   47
      Top             =   5760
      Width           =   495
   End
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   4
      Left            =   3600
      TabIndex        =   46
      Top             =   5760
      Width           =   495
   End
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   3
      Left            =   3360
      TabIndex        =   45
      Top             =   5760
      Width           =   495
   End
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   2
      Left            =   3120
      TabIndex        =   44
      Top             =   5760
      Width           =   495
   End
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   1
      Left            =   2880
      TabIndex        =   43
      Top             =   5760
      Width           =   495
   End
   Begin VB.ListBox lstA 
      Height          =   450
      Index           =   0
      Left            =   2640
      TabIndex        =   42
      Top             =   5760
      Width           =   495
   End
   Begin VB.ListBox lstFind6 
      Height          =   2010
      Left            =   1200
      TabIndex        =   41
      Top             =   3480
      Width           =   975
   End
   Begin VB.ListBox lstFind5 
      Height          =   2010
      Left            =   120
      TabIndex        =   40
      Top             =   3480
      Width           =   975
   End
   Begin VB.ListBox lstFind4 
      Height          =   4935
      Left            =   3600
      TabIndex        =   39
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   15
      ItemData        =   "frmBoggle.frx":038A
      Left            =   1440
      List            =   "frmBoggle.frx":0397
      TabIndex        =   38
      Top             =   7320
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   14
      ItemData        =   "frmBoggle.frx":03A7
      Left            =   960
      List            =   "frmBoggle.frx":03BA
      TabIndex        =   37
      Top             =   7320
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   13
      ItemData        =   "frmBoggle.frx":03D1
      Left            =   480
      List            =   "frmBoggle.frx":03E4
      TabIndex        =   36
      Top             =   7320
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   12
      ItemData        =   "frmBoggle.frx":03FA
      Left            =   0
      List            =   "frmBoggle.frx":0407
      TabIndex        =   35
      Top             =   7320
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   11
      ItemData        =   "frmBoggle.frx":0415
      Left            =   1440
      List            =   "frmBoggle.frx":0428
      TabIndex        =   34
      Top             =   6840
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   10
      ItemData        =   "frmBoggle.frx":043E
      Left            =   960
      List            =   "frmBoggle.frx":045A
      TabIndex        =   33
      Top             =   6840
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   9
      ItemData        =   "frmBoggle.frx":047A
      Left            =   480
      List            =   "frmBoggle.frx":0496
      TabIndex        =   32
      Top             =   6840
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   8
      ItemData        =   "frmBoggle.frx":04B6
      Left            =   0
      List            =   "frmBoggle.frx":04C9
      TabIndex        =   31
      Top             =   6840
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   7
      ItemData        =   "frmBoggle.frx":04DE
      Left            =   1440
      List            =   "frmBoggle.frx":04F1
      TabIndex        =   30
      Top             =   6360
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   6
      ItemData        =   "frmBoggle.frx":0506
      Left            =   960
      List            =   "frmBoggle.frx":0522
      TabIndex        =   29
      Top             =   6360
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   5
      ItemData        =   "frmBoggle.frx":0540
      Left            =   480
      List            =   "frmBoggle.frx":055C
      TabIndex        =   28
      Top             =   6360
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   4
      ItemData        =   "frmBoggle.frx":0579
      Left            =   0
      List            =   "frmBoggle.frx":058C
      TabIndex        =   27
      Top             =   6360
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   3
      ItemData        =   "frmBoggle.frx":059F
      Left            =   1440
      List            =   "frmBoggle.frx":05AC
      TabIndex        =   26
      Top             =   5880
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   2
      ItemData        =   "frmBoggle.frx":05B9
      Left            =   960
      List            =   "frmBoggle.frx":05CC
      TabIndex        =   25
      Top             =   5880
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   1
      ItemData        =   "frmBoggle.frx":05DF
      Left            =   480
      List            =   "frmBoggle.frx":05F2
      TabIndex        =   24
      Top             =   5880
      Width           =   495
   End
   Begin VB.ListBox lstLegal 
      Height          =   450
      Index           =   0
      ItemData        =   "frmBoggle.frx":0605
      Left            =   0
      List            =   "frmBoggle.frx":0612
      TabIndex        =   23
      Top             =   5880
      Width           =   495
   End
   Begin VB.ListBox lstFind3 
      Height          =   4935
      Left            =   2520
      TabIndex        =   22
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Words"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox lst6 
      Height          =   450
      Left            =   3360
      TabIndex        =   20
      Top             =   7320
      Width           =   1215
   End
   Begin VB.ListBox lst5 
      Height          =   450
      Left            =   2160
      TabIndex        =   19
      Top             =   7320
      Width           =   1215
   End
   Begin VB.ListBox lst4 
      Height          =   450
      Left            =   3360
      TabIndex        =   18
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ListBox lst3 
      Height          =   450
      Left            =   2160
      TabIndex        =   17
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   15
      Left            =   1800
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   14
      Left            =   1200
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   13
      Left            =   600
      TabIndex        =   13
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   12
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   10
      Left            =   1200
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   9
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   7
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Dictionary"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape shpA 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuAddToDict 
         Caption         =   "Add to Dictionary"
      End
   End
End
Attribute VB_Name = "frmBoggle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageFind Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long
Const WM_USER = &H400
Const LB_ERR = (-1)
Const LB_FINDSTRING = &H18F

Dim strTMP(6)
Dim Z, ZZ, ZZZ, ZZZZ, ZZZZZ, ZZZZZZ
Dim ZOOK0(23000), ZOOK1(23000), ZOOK2(23000)
Dim ZOOK3(23000), ZOOK4(23000), ZOOK5(23000)
Dim GetOut As Boolean

'this program will only find words that are between 3
'and 6 letters long...i did that to cut back on the required
'time for the search to run...right now, for me, this program
'takes a little under 2 minutes for a find(533MHz Celeron)
'...and that isn't too bad in my opinion...considering my
'first draft took well over 7 for each find...my conclusion:
'variables are always better that objects...and API can really
'make the difference sometimes.  I hope this program can help
'someone out with something they've been trying to do

'ps...all those lstLegal(x) on the form were/are used for
'the discover algorithm function
Private Sub cmdFind_Click()
If cmdFind.Caption = "Cancel" Then GetOut = True: Exit Sub
GetOut = False
Call FindAll
End Sub
Private Sub cmdMenu_Click()
frmBoggle.PopupMenu mnuFile
End Sub
Private Sub cmdReset_Click()
'clear the text boxes
GetOut = True
For Z = 0 To 15
    txtA(Z).Text = ""
Next Z
lstFind3.Clear
lstFind4.Clear
lstFind5.Clear
lstFind6.Clear
Label1.Visible = False
Label2.Visible = False
shpA.Visible = False
txtA(0).SetFocus
End Sub
Private Sub Form_Load()
Show
Dim wLIST3(50000) As String, wLIST4(50000) As String, wLIST5(50000) As String, wLIST6(50000) As String
Dim DD, ZiP3 As Integer, ZiP4 As Integer, ZiP5 As Integer, ZiP6 As Integer, LATT
Dim ZAD As Integer, AA(5), fPROG
ZAD = 0: ZiP3 = 0: ZiP4 = 0: ZiP5 = 0: ZiP6 = 0
shpA.Visible = True
'load up the dictionary file
LATT = FileLen(App.Path & "\dictionary.txt") 'get filesize of dict. file
Open App.Path & "\dictionary.txt" For Input As #1
    Do Until EOF(1)
        Input #1, DD
        DD = LCase(Trim(DD))
        If Len(DD) = 3 Then wLIST3(ZiP3) = DD: ZiP3 = ZiP3 + 1 'lst3.AddItem DD
        If Len(DD) = 4 Then wLIST4(ZiP4) = DD: ZiP4 = ZiP4 + 1 'lst4.AddItem DD
        If Len(DD) = 5 Then wLIST5(ZiP5) = DD: ZiP5 = ZiP5 + 1 'lst5.AddItem DD
        If Len(DD) = 6 Then wLIST6(ZiP6) = DD: ZiP6 = ZiP6 + 1 'lst6.AddItem DD
        fPROG = fPROG + Len(DD) + 2 'calculate how much of dict. file is open(add 2 for each carriage return after each word[from my observations])
        'show progress
        If fPROG Mod 30 = 0 Then
            shpA.Width = (fPROG / LATT) * Label1.Width
            Label2.Caption = "Loading Dictionary - " & Round((fPROG / LATT) * 100, 2) & "%"
        End If
        DoEvents
    Loop
Close #1
'make all the words into 4 strings, each word separated
'by a null character
Label2.Caption = "Compacting Dictionary - 0%"
strTMP(3) = Join(wLIST3, vbNullChar) '3 letter words
Label2.Caption = "Compacting Dictionary - 25%"
shpA.Width = 0.25 * Label1.Width
strTMP(4) = Join(wLIST4, vbNullChar) '4 letter words
Label2.Caption = "Compacting Dictionary - 50%"
shpA.Width = 0.5 * Label1.Width
strTMP(5) = Join(wLIST5, vbNullChar) '5 letter words
Label2.Caption = "Compacting Dictionary - 75%"
shpA.Width = 0.75 * Label1.Width
strTMP(6) = Join(wLIST6, vbNullChar) '6 letter words
Label2.Caption = "Compacting Dictionary - 100%"
shpA.Width = Label1.Width
ZAD = 0
'check to see if algorithm file is present
If LCase(Dir$(App.Path & "\algorithms.dat")) <> "algorithms.dat" Then
    'if algorithm file is messing, build a new one and exit sub
    Call DiscoverAlgorithms
    Exit Sub
End If
'load algorithm file
Open App.Path & "\algorithms.dat" For Input As #1
    Do Until EOF(1)
        Input #1, AA(0)
        Input #1, AA(1)
        Input #1, AA(2)
        Input #1, AA(3)
        Input #1, AA(4)
        Input #1, AA(5)
        
        ZOOK0(ZAD) = AA(0)
        ZOOK1(ZAD) = AA(1)
        ZOOK2(ZAD) = AA(2)
        ZOOK3(ZAD) = AA(3)
        ZOOK4(ZAD) = AA(4)
        ZOOK5(ZAD) = AA(5)
        ZAD = ZAD + 1
        
        'lstA(0).AddItem AA(0)
        'lstA(1).AddItem AA(1)
        'lstA(2).AddItem AA(2)
        'lstA(3).AddItem AA(3)
        'lstA(4).AddItem AA(4)
        'lstA(5).AddItem AA(5)
        'Label2.Caption = "Loading algorithms - " & Round((lstA(0).ListCount / 22671) * 100, 2) & "%"
        'shpA.Width = (lstA(0).ListCount / 22671) * Label1.Width
        
        'NOTE: i originally used listboxes to store the data
        'but it takes A LOT longer to load, and that is bad  ;o)
        
        DoEvents
        'show progress
        If ZAD Mod 30 = 0 Then
            Label2.Caption = "Loading algorithms - " & Round((ZAD / 22671) * 100, 2) & "%"
            shpA.Width = (ZAD / 22671) * Label1.Width
        End If
    Loop
Close #1
Label2.Visible = False
Label1.Visible = False
shpA.Visible = False
GetOut = False
cmdFind.Enabled = True
End Sub
Private Sub CheckList(NUMB As Integer, CAND As String)
'makes sure the found word isn't already in the found list.
'if it isn't, add it in
Dim RET
'QU correction...i replaced all the "QU" 's in the dictionary
'file with "æ"...the program will replace them too with the
'addtodict fucntion...the reason is that the program would
'consider "qu" to be one letter...ie it would look for the word
' "queen" in the list containing 4-letter words, thus it wouldn't
'find it, because queen has 5 letters...so i replaced all
'QU's with æ so the word "queen" woud be "æeen", and the program
'would find it in the 4-letter word list
If InStr(CAND, "æ") > 0 Then 'this just changes the æ character to a "qu" in the found list
    CAND = Left(CAND, InStr(CAND, "æ") - 1) & "qu" & Right(CAND, Len(CAND) - InStr(CAND, "æ"))
End If
Select Case NUMB
Case 3
    RET = SendMessageFind(lstFind3.hwnd, LB_FINDSTRING, 0, CAND)
    If RET = LB_ERR Then lstFind3.AddItem CAND
    Exit Sub
Case 4
    RET = SendMessageFind(lstFind4.hwnd, LB_FINDSTRING, 0, CAND)
    If RET = LB_ERR Then lstFind4.AddItem CAND
    Exit Sub
Case 5
    RET = SendMessageFind(lstFind5.hwnd, LB_FINDSTRING, 0, CAND)
    If RET = LB_ERR Then lstFind5.AddItem CAND
    Exit Sub
Case 6
    RET = SendMessageFind(lstFind6.hwnd, LB_FINDSTRING, 0, CAND)
    If RET = LB_ERR Then lstFind6.AddItem CAND
    Exit Sub
End Select
End Sub
Private Sub FindAll()
'the meat and potatoes of my program...this is where the
'real work is done...first it grabs the text box info and
'stores it to variables for faster access...then it uses
'the information in all the possible algorithms from my
'algorithm file(yes, there are 22672 of em)...quite a lot.
'it checks the formed words with the dictionary strings,
'and reports if it finds a match
'(by changing from my using the SendMessageFind api always to
'using the long strings, i dropped the seek-time from about 7
'minutes down to 3 minutes)
Dim LAST3 As String, LAST4 As String, LAST5 As String, WRD(6) As String
Dim tBOX(15) As String, STARTtime
GetOut = False
STARTtime = Now
cmdFind.Caption = "Cancel"
'grab data from the text boxes
For Z = 0 To 15
    tBOX(Z) = LCase(txtA(Z).Text)
    If LCase(txtA(Z).Text) = "qu" Then tBOX(Z) = "æ" 'QU correction
Next Z
Label1.Visible = True
Label2.Visible = True
shpA.Visible = True
'start finding the words
For Z = 0 To 22671
    On Error Resume Next
    'if user hit cancel, stop the function
    If GetOut = True Then cmdFind.Caption = "Find Words": Exit Sub
    '6 letter words
    WRD(6) = tBOX(ZOOK0(Z)) & tBOX(ZOOK1(Z)) & tBOX(ZOOK2(Z)) & tBOX(ZOOK3(Z)) & tBOX(ZOOK4(Z)) & tBOX(ZOOK5(Z))
    If InStr(strTMP(6), WRD(6)) > 0 Then 'is the formed word contained in the dictionary string?
         Call CheckList(6, WRD(6)) 'if yes, then see if it was found previously
    End If
    '5 letter
    WRD(5) = tBOX(ZOOK0(Z)) & tBOX(ZOOK1(Z)) & tBOX(ZOOK2(Z)) & tBOX(ZOOK3(Z)) & tBOX(ZOOK4(Z))
    If WRD(5) <> LAST5 Then
        If InStr(strTMP(5), WRD(5)) > 0 Then
            Call CheckList(5, WRD(5))
        End If
    End If
    '4 letter
    WRD(4) = tBOX(ZOOK0(Z)) & tBOX(ZOOK1(Z)) & tBOX(ZOOK2(Z)) & tBOX(ZOOK3(Z))
    If WRD(4) <> LAST4 Then
        If InStr(strTMP(4), WRD(4)) > 0 Then
            Call CheckList(4, WRD(4))
        End If
    End If
    '3 letter
    WRD(3) = tBOX(ZOOK0(Z)) & tBOX(ZOOK1(Z)) & tBOX(ZOOK2(Z))
    If WRD(3) <> LAST3 Then
        If InStr(strTMP(3), WRD(3)) > 0 Then
            Call CheckList(3, WRD(3))
        End If
    End If
    'to prevent redundancy and add a little speed(dropped
    'a whole minute of scan-time off for me)
    LAST3 = WRD(3)
    LAST4 = WRD(4)
    LAST5 = WRD(5)
    'show progress
    If Z Mod 50 = 0 Then
        Label2.Caption = "Finding Words - " & Round((Z / 22671) * 100, 2) & "%"
        shpA.Width = (Z / 22671) * Label1.Width
    End If
    DoEvents
Next Z
'we're done...hide the progress meter and display how long
'the search took
STARTtime = DateDiff("s", STARTtime, Now)
Label2.Caption = "Scan took: " & STARTtime \ 60 & " minutes and " & STARTtime Mod 60 & " seconds."
Label1.Visible = False
'Label2.Visible = False
shpA.Visible = False
cmdFind.Caption = "Find Words"
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub DiscoverAlgorithms()
'this is actually the code i used to find all possible jumps
'from one letter to another...i then used a different code
'to pull out double jumps(ex.  1 to 0 to 1...since you can
'only use each letter once)...the code to remaove all the double
'jumps is in frmboggletest.  i included the algorithm finder
'in case the user accidently deletes the algorithm file, the program
'will detect that its missing and rebuild it.
Dim DONEamt
Dim POT(5)
If MsgBox("You are missing the algorithm file.  Would you like it rebuilt?" & vbCrLf & "(it is required for this program to run)", vbYesNo, "Missing Algorithm File") = vbNo Then End
shpA.Visible = True
Label1.Visible = True
Label2.Visible = True
Open App.Path & "\temp.txt" For Output As #1
For Z = 0 To 15
    POT(0) = Z
    For ZZ = 0 To lstLegal(POT(0)).ListCount - 1
        POT(1) = lstLegal(POT(0)).List(ZZ)
        For ZZZ = 0 To lstLegal(POT(1)).ListCount - 1
            DONEamt = DONEamt + 1
            'show status
            Label2.Caption = "Processing temporary algorithms - " & Round((DONEamt / 492) * 100, 2) & "%"
            shpA.Width = (DONEamt / 492) * Label1.Width
            POT(2) = lstLegal(POT(1)).List(ZZZ)
            For ZZZZ = 0 To lstLegal(POT(2)).ListCount - 1
                POT(3) = lstLegal(POT(2)).List(ZZZZ)
                For ZZZZZ = 0 To lstLegal(POT(3)).ListCount - 1
                    POT(4) = lstLegal(POT(3)).List(ZZZZZ)
                    For ZZZZZZ = 0 To lstLegal(POT(4)).ListCount - 1
                        POT(5) = lstLegal(POT(4)).List(ZZZZZZ)
                        Print #1, POT(0) & "," & POT(1) & "," & POT(2) & "," & POT(3) & "," & POT(4) & "," & POT(5)
                        DoEvents
                    Next ZZZZZZ
                Next ZZZZZ
            Next ZZZZ
        Next ZZZ
    Next ZZ
Next Z
Close #1
frmboggletest.Show
End Sub
Private Sub mnuAbout_Click()
frmAboutAll.Show
End Sub
Private Sub mnuAddToDict_Click()
Call ADDtoDICT
End Sub
Private Sub txtA_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'auto-move to next textbox when something is inputted
If LCase(txtA(Index).Text) = "q" Then txtA(Index).Text = "qu"
If Index < 15 Then txtA(Index + 1).SetFocus
End Sub
Private Sub ADDtoDICT()
'lets you add more words to the dictionary file...it makes sure
'that no duplicate words are added also
MsgBox "Choose the file you want to add to the dictionary", vbOKOnly, "Add to Dictionary"
CD.Filter = "Text files(*.txt)|*.txt"
CD.ShowOpen
If CD.FileName = "" Then Exit Sub
Dim TIZE, TESTVAL
TIZE = FreeFile
Open CD.FileName For Input As #TIZE
    Do Until EOF(TIZE)
        Input #TIZE, TESTVAL
        TESTVAL = LCase(Trim(TESTVAL))
        If InStr(textval, "qu") > 0 Then 'QU correction(see CheckList sub)
            TESTVAL = Left(TESTVAL, InStr(TESTVAL, "qu") - 1) & "æ" & Right(TESTVAL, Len(TESTVAL) - InStr(TESTVAL, "qu") - 1)
        End If
        If Len(TESTVAL) = 3 Then
            If InStr(strTMP(3), TESTVAL) > 0 Then
                lst4.AddItem TESTVAL
            Else
                lst3.AddItem TESTVAL
            End If
        End If
        If Len(TESTVAL) = 4 Then
            If InStr(strTMP(4), TESTVAL) > 0 Then
                lst4.AddItem TESTVAL
            Else
                lst3.AddItem TESTVAL
            End If
        End If
        If Len(TESTVAL) = 5 Then
            If InStr(strTMP(5), TESTVAL) > 0 Then
                lst4.AddItem TESTVAL
            Else
                lst3.AddItem TESTVAL
            End If
        End If
        If Len(TESTVAL) = 6 Then
            If InStr(strTMP(6), TESTVAL) > 0 Then
                lst4.AddItem TESTVAL
            Else
                lst3.AddItem TESTVAL
            End If
        End If
        DoEvents
    Loop
Close #TIZE
TIZE = FreeFile
Open App.Path & "\dictionary.txt" For Append As #TIZE
    For ZZZ = 0 To lst3.ListCount - 1
        Print #TIZE, lst3.List(ZZZ)
    Next ZZZ
Close #TIZE
MsgBox lst3.ListCount & " files were added to the dictionary." & vbCrLf & lst4.ListCount & " duplicate words were found and not added." & vbCrLf & vbCrLf & "Changes will take effect next time you open the program.", vbOKOnly, "Words Added"
lst3.Clear
lst4.Clear
End Sub
