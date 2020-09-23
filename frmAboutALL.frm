VERSION 5.00
Begin VB.Form frmAboutAll 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmAboutALL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScroll 
      Interval        =   100
      Left            =   120
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1800
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   0
      X1              =   1080
      X2              =   840
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   10
      X1              =   3480
      X2              =   3240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   9
      X1              =   3000
      X2              =   3240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   8
      X1              =   2760
      X2              =   3000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   7
      X1              =   2520
      X2              =   2760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   6
      X1              =   2280
      X2              =   2520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   5
      X1              =   2040
      X2              =   2280
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   4
      X1              =   1800
      X2              =   2040
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   3
      X1              =   1560
      X2              =   1800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   1
      X1              =   1080
      X2              =   1320
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LB 
      BorderWidth     =   2
      Index           =   2
      X1              =   1560
      X2              =   1320
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   0
      X1              =   1080
      X2              =   840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   10
      X1              =   3480
      X2              =   3240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   9
      X1              =   3000
      X2              =   3240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   8
      X1              =   2760
      X2              =   3000
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LS 
      BorderWidth     =   2
      Index           =   0
      X1              =   3480
      X2              =   3480
      Y1              =   2400
      Y2              =   2640
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   7
      X1              =   2520
      X2              =   2760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   6
      X1              =   2280
      X2              =   2520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   5
      X1              =   2040
      X2              =   2280
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   4
      X1              =   1800
      X2              =   2040
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   3
      X1              =   1560
      X2              =   1800
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   1
      X1              =   1080
      X2              =   1320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   2
      X1              =   1560
      X2              =   1320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LS 
      BorderWidth     =   2
      Index           =   1
      X1              =   840
      X2              =   840
      Y1              =   2640
      Y2              =   2400
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   25
      Left            =   3960
      TabIndex        =   29
      Top             =   2400
      Width           =   150
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   24
      Left            =   3840
      TabIndex        =   28
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   23
      Left            =   3720
      TabIndex        =   27
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   22
      Left            =   3600
      TabIndex        =   26
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   3480
      TabIndex        =   25
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   20
      Left            =   3360
      TabIndex        =   24
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   19
      Left            =   3240
      TabIndex        =   23
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   3120
      TabIndex        =   22
      Top             =   2400
      Width           =   60
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   17
      Left            =   3000
      TabIndex        =   21
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   16
      Left            =   2880
      TabIndex        =   20
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   15
      Left            =   2760
      TabIndex        =   19
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   14
      Left            =   2640
      TabIndex        =   18
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   2520
      TabIndex        =   17
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   2400
      TabIndex        =   16
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   11
      Left            =   2280
      TabIndex        =   15
      Top             =   2400
      Width           =   60
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   10
      Left            =   2160
      TabIndex        =   14
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   9
      Left            =   2040
      TabIndex        =   13
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   8
      Left            =   1920
      TabIndex        =   12
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   1800
      TabIndex        =   11
      Top             =   2400
      Width           =   60
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   6
      Left            =   1680
      TabIndex        =   10
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label LX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   2400
      Width           =   120
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   990
      Left            =   550
      Picture         =   "frmAboutALL.frx":030A
      Top             =   1005
      Width           =   3645
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   990
      Left            =   575
      Picture         =   "frmAboutALL.frx":C84C
      Top             =   960
      Width           =   3645
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   45
      X1              =   360
      X2              =   0
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   44
      X1              =   720
      X2              =   360
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   43
      X1              =   1080
      X2              =   720
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   42
      X1              =   1440
      X2              =   1080
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   41
      X1              =   1800
      X2              =   1440
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   40
      X1              =   2160
      X2              =   1800
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   39
      X1              =   2520
      X2              =   2160
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   38
      X1              =   2880
      X2              =   2520
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   37
      X1              =   3240
      X2              =   2880
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   36
      X1              =   3600
      X2              =   3240
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   35
      X1              =   3960
      X2              =   3600
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   34
      X1              =   4320
      X2              =   3960
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   33
      X1              =   4680
      X2              =   4320
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   32
      X1              =   4440
      X2              =   4440
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   31
      X1              =   4440
      X2              =   4440
      Y1              =   720
      Y2              =   360
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   30
      X1              =   4440
      X2              =   4440
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   29
      X1              =   4440
      X2              =   4440
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   28
      X1              =   4440
      X2              =   4440
      Y1              =   1800
      Y2              =   1440
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   27
      X1              =   4440
      X2              =   4440
      Y1              =   2160
      Y2              =   1800
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   26
      X1              =   4440
      X2              =   4440
      Y1              =   2520
      Y2              =   2160
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   25
      X1              =   4440
      X2              =   4440
      Y1              =   2880
      Y2              =   2520
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   24
      X1              =   4440
      X2              =   4440
      Y1              =   3240
      Y2              =   2880
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   23
      X1              =   4440
      X2              =   4440
      Y1              =   3600
      Y2              =   3240
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   22
      X1              =   4680
      X2              =   4320
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   21
      X1              =   4320
      X2              =   3960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   20
      X1              =   3960
      X2              =   3600
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   19
      X1              =   3600
      X2              =   3240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   18
      X1              =   3240
      X2              =   2880
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   17
      X1              =   2880
      X2              =   2520
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   16
      X1              =   2520
      X2              =   2160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   15
      X1              =   2160
      X2              =   1800
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   14
      X1              =   1800
      X2              =   1440
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   13
      X1              =   1440
      X2              =   1080
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   12
      X1              =   1080
      X2              =   720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   11
      X1              =   720
      X2              =   360
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   10
      X1              =   360
      X2              =   0
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   9
      X1              =   15
      X2              =   15
      Y1              =   3600
      Y2              =   3240
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   8
      X1              =   15
      X2              =   15
      Y1              =   3240
      Y2              =   2880
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   7
      X1              =   15
      X2              =   15
      Y1              =   2880
      Y2              =   2520
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   6
      X1              =   15
      X2              =   15
      Y1              =   2520
      Y2              =   2160
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   5
      X1              =   15
      X2              =   15
      Y1              =   2160
      Y2              =   1800
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   4
      X1              =   15
      X2              =   15
      Y1              =   1800
      Y2              =   1440
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   3
      X1              =   15
      X2              =   15
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   2
      X1              =   15
      X2              =   15
      Y1              =   1080
      Y2              =   720
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   1
      X1              =   15
      X2              =   15
      Y1              =   720
      Y2              =   360
   End
   Begin VB.Line PR 
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   15
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   2295
      X2              =   2495
      Y1              =   3525
      Y2              =   3525
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   2100
      X2              =   2300
      Y1              =   3525
      Y2              =   3525
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   1905
      X2              =   2105
      Y1              =   3525
      Y2              =   3525
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   1905
      X2              =   1905
      Y1              =   3315
      Y2              =   3515
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   1905
      X2              =   1905
      Y1              =   3320
      Y2              =   3120
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   1905
      X2              =   2105
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   2300
      X2              =   2100
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   2295
      X2              =   2495
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   2505
      X2              =   2505
      Y1              =   3120
      Y2              =   3320
   End
   Begin VB.Line LN 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   2505
      X2              =   2505
      Y1              =   3515
      Y2              =   3315
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   2
      Top             =   3120
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Production"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4575
   End
End
Attribute VB_Name = "frmAboutAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim D, EE(46), A, B, C, FF
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, ByVal _
lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
D = 0
For Z = 1800 To 1200 Step -15
    D = D + 1
    Line (400, 1800)-(4200, Z), EE(D)
    Line (4200, 1200)-(400, 3000 - Z), EE(D)
    Line (400, 1200)-(4200, 3000 - Z), EE(D)
    Line (4200, 1800)-(400, Z), EE(D)
Next Z
If A = 250 Then C = 1
If A = 10 Then C = 0
If C = 0 Then A = A + 10: B = B - 10
If C = 1 Then A = A - 10: B = B + 10
EE(1) = RGB(A, 0, B)
For Z = 46 To 2 Step -1
    EE(Z) = EE(Z - 1)
Next Z
For Z = 0 To 9
    LN(Z).BorderColor = EE(2 * Z + 1)
Next Z
Label5.ForeColor = EE(21)
For Z = 0 To 45
    PR(Z).BorderColor = EE(Z + 1)
Next Z
LX(12).ForeColor = EE(5)
For Z = 25 To 13 Step -1
    LX(Z).ForeColor = LX(Z - 1).ForeColor
Next Z
For Z = 0 To 11
    LX(Z).ForeColor = LX(Z + 1).ForeColor
Next Z
For Z = 0 To 4
    LB(Z).BorderColor = LX(10 - 2 * Z).ForeColor
    LT(Z).BorderColor = LX(10 - 2 * Z).ForeColor
Next Z
LB(5).BorderColor = LX(24).ForeColor
LT(5).BorderColor = LX(24).ForeColor
For Z = 5 To 10
    LB(Z).BorderColor = LX(2 * Z - 8).ForeColor
    LT(Z).BorderColor = LX(2 * Z - 8).ForeColor
Next Z
LS(0).BorderColor = LX(12).ForeColor
LS(1).BorderColor = LX(12).ForeColor
End Sub
Private Sub Form_Load()
For Z = 1 To 25
    LX(Z).Left = LX(Z - 1).Left + LX(Z - 1).Width
Next Z
A = 0
B = 250
C = 0
For Z = 0 To 46
    EE(Z) = 0
Next Z
For Z = 0 To 10
    LT(Z).Visible = False
    LB(Z).Visible = False
Next Z
LS(0).Visible = False
LS(1).Visible = False
FF = "     This program was created by Ryan DuPont, aka The StainMaster.  I hope you like it.  Send any questions or comments to NoNamedGuy@hotmail.com   ¤¤¤"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call HideSite
For Z = 0 To 9
    LN(Z).Visible = False
Next Z
End Sub
Private Sub Label3_Click()
Call HideSite
End Sub
Private Sub Label5_Click()
frmAboutAll.Cls
Unload frmAboutAll
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For Z = 0 To 9
    LN(Z).Visible = True
Next Z
End Sub

Private Sub LX_Click(Index As Integer)
Call ShellExecute(0&, vbNullString, "http://lutzlutz.tripod.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub LX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For Z = 0 To 10
    LT(Z).Visible = True
    LB(Z).Visible = True
Next Z
LS(0).Visible = True
LS(1).Visible = True
End Sub

Private Sub Timer1_Timer()
Call Command1_Click
End Sub

Private Sub tmrScroll_Timer()
FF = Mid(FF, 2, Len(FF) - 1) & Mid(FF, 1, 1)
Label6.Caption = FF
End Sub
Private Sub HideSite()
For Z = 0 To 10
    LT(Z).Visible = False
    LB(Z).Visible = False
Next Z
LS(0).Visible = False
LS(1).Visible = False
End Sub
