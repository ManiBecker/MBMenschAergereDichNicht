VERSION 5.00
Begin VB.Form lblStart 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MBMenschAergereDichNicht"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MBMenschAergereDichNicht"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox picWuerfel 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4560
      ScaleHeight     =   1095
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   1560
      Width           =   855
      Begin VB.Shape shpWuerfelAuge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H0000C0C0&
         Height          =   135
         Index           =   0
         Left            =   360
         Shape           =   3  'Kreis
         Top             =   360
         Width           =   135
      End
      Begin VB.Shape shpWuerfelAuge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H0000C0C0&
         Height          =   135
         Index           =   1
         Left            =   120
         Shape           =   3  'Kreis
         Top             =   120
         Width           =   135
      End
      Begin VB.Shape shpWuerfelAuge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H0000C0C0&
         Height          =   135
         Index           =   2
         Left            =   600
         Shape           =   3  'Kreis
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape shpWuerfelAuge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H0000C0C0&
         Height          =   135
         Index           =   3
         Left            =   600
         Shape           =   3  'Kreis
         Top             =   120
         Width           =   135
      End
      Begin VB.Shape shpWuerfelAuge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H0000C0C0&
         Height          =   135
         Index           =   4
         Left            =   120
         Shape           =   3  'Kreis
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape shpWuerfelAuge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H0000C0C0&
         Height          =   135
         Index           =   5
         Left            =   120
         Shape           =   3  'Kreis
         Top             =   360
         Width           =   135
      End
      Begin VB.Shape shpWuerfelAuge 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H0000C0C0&
         Height          =   135
         Index           =   6
         Left            =   600
         Shape           =   3  'Kreis
         Top             =   360
         Width           =   135
      End
      Begin VB.Shape shpWuerfelPaper 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Undurchsichtig
         Height          =   855
         Left            =   0
         Shape           =   3  'Kreis
         Top             =   0
         Width           =   855
      End
      Begin VB.Shape shpWuerfelBorder 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Undurchsichtig
         Height          =   855
         Left            =   0
         Shape           =   4  'Gerundetes Rechteck
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblWuerfel 
         Alignment       =   2  'Zentriert
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   240
      Top             =   1680
   End
   Begin VB.Label lblHint 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "lblHint"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7080
      Width           =   6975
   End
   Begin VB.Label lblPlatzierung 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblPlatzierung 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblPlatzierung 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblPlatzierung 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   0
      Left            =   5640
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   375
      Index           =   15
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   375
      Index           =   14
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   375
      Index           =   13
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   375
      Index           =   12
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   11
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   10
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   9
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   8
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   375
      Index           =   7
      Left            =   6360
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   375
      Index           =   6
      Left            =   5760
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   375
      Index           =   5
      Left            =   6360
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   375
      Index           =   4
      Left            =   5760
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      Height          =   375
      Index           =   3
      Left            =   6240
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      Height          =   375
      Index           =   2
      Left            =   5640
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      Height          =   375
      Index           =   1
      Left            =   6240
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape shpSpieler 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      Height          =   375
      Index           =   0
      Left            =   5640
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   435
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   240
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   0
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   2
      Left            =   2760
      TabIndex        =   4
      Top             =   6240
      Width           =   240
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   435
      Index           =   1
      Left            =   6360
      TabIndex        =   3
      Top             =   3840
      Width           =   240
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   55
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   54
      Left            =   2040
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   53
      Left            =   1440
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   52
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   51
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   50
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   49
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   5040
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   48
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   47
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   46
      Left            =   4440
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   45
      Left            =   5040
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   44
      Left            =   5640
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   43
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   42
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   41
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   40
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   15
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   14
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   13
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   12
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   11
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   10
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   9
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   8
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   7
      Left            =   6360
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   6
      Left            =   5760
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   5
      Left            =   6360
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   4
      Left            =   5760
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   3
      Left            =   6240
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   2
      Left            =   5640
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   1
      Left            =   6240
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape shpStart 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   0
      Left            =   5640
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   39
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   38
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   37
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   36
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   35
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   34
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   33
      Left            =   2040
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   32
      Left            =   1440
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   31
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   30
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   29
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   28
      Left            =   240
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   27
      Left            =   840
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   26
      Left            =   1440
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   25
      Left            =   2040
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   24
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   23
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   22
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   5040
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   21
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   20
      Left            =   2640
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   19
      Left            =   3240
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   18
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   17
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   16
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   5040
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   15
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   14
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   13
      Left            =   4440
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   12
      Left            =   5040
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   11
      Left            =   5640
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   10
      Left            =   6240
      Shape           =   3  'Kreis
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   9
      Left            =   6240
      Shape           =   3  'Kreis
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   8
      Left            =   6240
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   7
      Left            =   5640
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   6
      Left            =   5040
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   5
      Left            =   4440
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   4
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   3
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   2
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpFeld 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Undurchsichtig
      Height          =   495
      Index           =   1
      Left            =   3840
      Shape           =   3  'Kreis
      Top             =   840
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   10
      X1              =   2880
      X2              =   2880
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   9
      X1              =   2880
      X2              =   480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   8
      X1              =   480
      X2              =   480
      Y1              =   2880
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   7
      X1              =   480
      X2              =   2880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   6
      X1              =   2880
      X2              =   2880
      Y1              =   4080
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   5
      X1              =   2880
      X2              =   4080
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   4
      X1              =   4080
      X2              =   4080
      Y1              =   4080
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   3
      X1              =   6480
      X2              =   4080
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   6480
      X2              =   6480
      Y1              =   2880
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   4080
      X2              =   6480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   4080
      X2              =   4080
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   11
      X1              =   2880
      X2              =   4080
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "lblStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////
'//
'// Projekt : MBMenschAergereDichNicht
'// Sprache : Basic
'// Compiler: MS Visual Basic 6.0
'// Autor   : Manfred Becker
'// E-Mail  : mani.becker@web.de
'// Url     : http://manib.ma.funpic.de
'// Modul   : frmMain
'// Version : 1.00
'// Datum   : 04.09.2008
'//
'///////////////////////////////////////////////////////////////////////////

Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim dx As Integer
Dim dy As Integer
Dim player As Integer
Dim startIndex(4) As Integer
Dim zielIndex(4) As Integer
Dim spielerIndex(4) As Integer
Dim spielerFarbe(4) As String
Dim wuerfelPosTop(4) As Integer
Dim wuerfelPosLeft(4) As Integer
Dim anzahlWuerfe(4) As Integer
Dim streckenIndex(4, 44) As Integer
Dim platzierung(4) As Integer
Dim platz As Integer

Private Sub ZeigeHinweis(hint As String)
    lblHint.Caption = hint
    VBSleep (1000)
End Sub

Private Sub InitializeGame()
    Dim i As Integer
    Dim j As Integer
    Dim index As Integer
    Randomize Timer
    
    platz = 0
    
    For i = 0 To 3
        startIndex(i) = i * 4
        zielIndex(i) = 40 + i * 4
        spielerIndex(i) = i * 4
        anzahlWuerfe(i) = 3
        platzierung(i) = 0
        lblPlatzierung(i).Visible = False
        For j = 0 To 43
            If j < 40 Then
                index = i * 10 + j
                If index >= 40 Then
                    index = index - 40
                End If
            Else
                index = zielIndex(i) + (j - 40)
            End If
            streckenIndex(i, j) = index
        Next j
    Next i
    
    dx = (shpStart(0).Width - shpSpieler(0).Width) / 2
    dy = (shpStart(0).Height - shpSpieler(0).Height) / 2
    
    For i = 0 To shpSpieler.Count - 1
        shpSpieler(i).Top = shpStart(i).Top + dy
        shpSpieler(i).Left = shpStart(i).Left + dx
        shpSpieler(i).Tag = ""
    Next i
    
    For i = 0 To shpStart.Count - 1
        shpStart(i).Tag = i
    Next i
    For i = 0 To shpFeld.Count - 1
        shpFeld(i).Tag = ""
    Next i
    
    wuerfelPosTop(0) = 1560
    wuerfelPosLeft(0) = 4560
    wuerfelPosTop(1) = 4560
    wuerfelPosLeft(1) = 4560
    wuerfelPosTop(2) = 4560
    wuerfelPosLeft(2) = 1560
    wuerfelPosTop(3) = 1560
    wuerfelPosLeft(3) = 1560
    spielerFarbe(0) = "Gelb"
    spielerFarbe(1) = "Grün"
    spielerFarbe(2) = "Rot"
    spielerFarbe(3) = "Blau"
    
    player = 0
    Call NextPlayer
    
    Timer1.Enabled = False
    
End Sub

Private Sub StartGame()
    Call InitializeGame
    player = ErmittleSpielBeginner
    ZeigeHinweis "Spiel läuft..."
    Timer1.Enabled = True
End Sub


Private Function ErmittleSpielBeginner() As Integer
    Dim i As Integer
    Dim z As Integer
    Dim s As Integer
    Dim n As Integer
    
    ZeigeHinweis "Ermittle Spieler der beginnt..."
    
    For i = 0 To 3
        platzierung(i) = 0
    Next i
    
    Do
        z = 0
        s = -1
        n = 0
        For i = 0 To 3
            If platzierung(i) = 0 Then
                picWuerfel.Top = wuerfelPosTop(i)
                picWuerfel.Left = wuerfelPosLeft(i)
                platzierung(i) = Wuerfle()
                If z < platzierung(i) Then
                    z = platzierung(i)
                    s = i
                End If
                VBSleep (1000)
            End If
        Next i
        For i = 0 To 3
            If z > platzierung(i) Then
                platzierung(i) = -1
                n = n + 1
            Else
                platzierung(i) = 0
            End If
        Next i
    Loop While n < 3
    
    
    For i = 0 To 3
        platzierung(i) = 0
    Next i

    ZeigeHinweis spielerFarbe(s) & " beginnt!"
    ErmittleSpielBeginner = s
End Function

Private Sub NextPlayer()

    If platz = 4 Then Exit Sub

    If anzahlWuerfe(player) > 0 Then
        anzahlWuerfe(player) = anzahlWuerfe(player) - 1
    Else
        Do
            player = player + 1
            If player > 3 Then player = 0
        Loop Until platzierung(player) = 0
        
        If IstSpielerAufSpielfeld(player) Then
            anzahlWuerfe(player) = 1
        Else
            anzahlWuerfe(player) = 2
        End If
    End If
    
    picWuerfel.Top = wuerfelPosTop(player)
    picWuerfel.Left = wuerfelPosLeft(player)
    Call ZeichneWuerfel(0)
    
End Sub

Private Function IstStartFeldFrei(p As Integer) As Boolean
    Dim bRet As Boolean
    bRet = False
    
    If shpFeld(streckenIndex(p, 0)).Tag = "" Then bRet = True
    
    IstStartFeldFrei = bRet
End Function

Private Function IstSpielerZuHause(p As Integer) As Boolean
    Dim bRet As Boolean
    Dim i As Integer
    bRet = False
    
    For i = 0 To 3
        If shpStart(startIndex(p) + i).Tag <> "" Then
            bRet = True
            Exit For
        End If
    Next i
    
    IstSpielerZuHause = bRet
End Function

Private Sub SetzeSpielerAufSpielfeld(p As Integer)
    Dim i As Integer
    Dim n As Integer
    
    For i = startIndex(p) To startIndex(p) + 3
        If shpStart(i).Tag <> "" Then
            n = shpStart(i).Tag
            shpSpieler(n).Left = shpFeld(streckenIndex(p, 0)).Left + dx
            shpSpieler(n).Top = shpFeld(streckenIndex(p, 0)).Top + dy
            shpSpieler(n).Tag = 0
            shpStart(i).Tag = ""
            shpFeld(streckenIndex(p, 0)).Tag = i
            anzahlWuerfe(p) = 0
            Exit For
        End If
    Next i
End Sub

Private Function IstSpielerAufSpielfeld(p As Integer) As Boolean
    Dim i As Integer
    Dim bRet As Boolean
    bRet = False
    
    For i = spielerIndex(p) To spielerIndex(p) + 3
        If shpSpieler(i).Tag <> "" Then
            bRet = True
            Exit For
        End If
    Next i
    
    IstSpielerAufSpielfeld = bRet
End Function

Private Sub ZieheSieler(p As Integer, i As Integer, s As Integer)
    Dim f As Integer
    Dim z As Integer
    Dim g As Integer
    
    f = streckenIndex(p, shpSpieler(i).Tag)
    z = streckenIndex(p, shpSpieler(i).Tag + s)
    
    'das alte Feld wird frei
    shpFeld(f).Tag = ""
    
    'Gegner?
    If shpFeld(z).Tag <> "" Then
        g = shpFeld(z).Tag

        'Gegner schlagen
        shpSpieler(g).Left = shpStart(g).Left + dx
        shpSpieler(g).Top = shpStart(g).Top + dy
        shpSpieler(g).Tag = ""

        shpStart(g).Tag = g
    End If
    
    'das neue Feld wird besetzt
    shpFeld(z).Tag = i
    
    'Zielfeld sichern
    shpSpieler(i).Tag = shpSpieler(i).Tag + s
    
    'Figur umsetzen
    shpSpieler(i).Left = shpFeld(z).Left + dx
    shpSpieler(i).Top = shpFeld(z).Top + dy
    
    'Fertig?
    Dim fertig As Boolean
    fertig = True
    For f = 40 To 43
        z = streckenIndex(p, f)
        If shpFeld(z).Tag = "" Then
            fertig = False
            Exit For
        End If
    Next f
    If fertig Then
        platz = platz + 1
        lblPlatzierung(p).Caption = platz
        lblPlatzierung(p).Visible = True
        platzierung(p) = platz
        ZeigeHinweis "Spieler " & spielerFarbe(p) & " ist auf Platz " & platz
    End If
End Sub

Private Sub BewegeSpieler(p As Integer, s As Integer)
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim f As Integer
    Dim z As Integer
    Dim fi As Integer
    Dim derWegIstFrei As Boolean
    Dim ziel As String
    
    
    If Not IstSpielerAufSpielfeld(p) Then Exit Sub
    
    'Ist das Startfeld besetzt?
    f = streckenIndex(p, 0)
    If shpFeld(f).Tag <> "" Then
        'Steht dort der eigene Spielstein?
        i = shpFeld(f).Tag
        If i >= startIndex(p) And i <= startIndex(p) + 3 Then
            'Kann gezogen werden?
            z = streckenIndex(p, shpSpieler(i).Tag + s)
            
            If shpFeld(z).Tag = "" Then 'Or shpFeld(z).Tag <> n Then
                Call ZieheSieler(p, i, s)
                Exit Sub
            End If
        End If
    End If
    
    'Gibt es einen Spielstein im Ziel, der noch bewegt werden muss?
    If s < 4 Then
        ziel = ""
        For i = 43 To 40 Step -1
            z = streckenIndex(p, i)
            If shpFeld(z).Tag <> "" Then
                ziel = ziel & "X"
            Else
                ziel = ziel & " "
            End If
        Next i
        If ziel <> "    " And ziel <> "XXXX" Then
            i = -1
            If (s = 3 And ziel = "   X") Then
                i = shpFeld(streckenIndex(p, 40)).Tag
            ElseIf (s = 2 And InStr(1, ziel, "  X") > 0) Then
                z = 2 - InStr(1, ziel, "  X")
                i = shpFeld(streckenIndex(p, 40) + z).Tag
            ElseIf (s = 1 And InStr(1, ziel, " X") > 0) Then
                z = 3 - InStr(1, ziel, " X")
                i = shpFeld(streckenIndex(p, 40) + z).Tag
            End If
               
            If i > -1 Then
                Call ZieheSieler(p, i, s)
                Exit Sub
            End If
        End If
    End If
    
    
    'Kann mit einem Spielstein gezogen werden?
    For i = spielerIndex(p) To spielerIndex(p) + 3
        'Befindet sich der Spielstein auf dem Spielfeld?
        If shpSpieler(i).Tag <> "" Then
            f = streckenIndex(p, shpSpieler(i).Tag)
            
            j = shpSpieler(i).Tag + s
            If j <= 43 Then
            
                'Zielfeld berechnen
                z = streckenIndex(p, shpSpieler(i).Tag + s)
                
                'bereits im Ziel?
                If j >= 40 Then
                    'Spielstein befindet sich im Ziel
                    'kann gezogen werden?
                    derWegIstFrei = True
                    For n = 40 To j
                        If shpFeld(streckenIndex(p, n)).Tag <> "" Then
                            derWegIstFrei = False
                            Exit For
                        End If
                    Next n
                    
                    If derWegIstFrei Then
                        Call ZieheSieler(p, i, s)
                        Exit Sub
                    End If
                Else
                    'Spielstein befindet sich auf dem Feld
                    If shpFeld(z).Tag <> "" Then
                        'Zielfeld ist besetzt
                        'Kann gezogen werden?
                        n = shpFeld(z).Tag
                        If n < startIndex(p) Or n > startIndex(p) + 3 Then
                            'Schlage Gegner
                            Call ZieheSieler(p, i, s)
                            Exit Sub
                        End If
                    Else
                        'Zielfeld ist leer
                        Call ZieheSieler(p, i, s)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Function NextStep() As Boolean
    Dim GameOver As Boolean
    Dim w As Integer
    
    GameOver = True
    w = Wuerfle
    
    
    If w = 6 Then
        If IstSpielerZuHause(player) Then
            If IstStartFeldFrei(player) Then
                Call SetzeSpielerAufSpielfeld(player)
            Else
                Call BewegeSpieler(player, w)
            End If
        Else
            Call BewegeSpieler(player, w)
        End If
    ElseIf w >= 1 And w < 6 Then
        If IstSpielerAufSpielfeld(player) Then
            Call BewegeSpieler(player, w)
            anzahlWuerfe(player) = 0
        End If
        
        Call NextPlayer
    Else
        MsgBox "Falsche Augenzahl: " & Str(w)
    End If
    
    'Spiel beendet?
    If platz = 4 Then
        GameOver = False
    End If
    
    
    NextStep = GameOver
End Function

Private Function Wuerfle() As Integer
    Dim i As Integer
    
    i = Int(Rnd(1) * 6) + 1
    
    Call ZeichneWuerfel(i)
    Call VBSleep(20)
    
    Wuerfle = i
End Function

Private Sub ZeichneWuerfel(i As Integer)
    If i >= 1 And i <= 6 Then
        lblWuerfel.Caption = i
    Else
        lblWuerfel.Caption = "?"
    End If
    
    If i = 1 Or i = 3 Or i = 5 Then
        shpWuerfelAuge(0).Visible = True
    Else
        shpWuerfelAuge(0).Visible = False
    End If
    If i = 2 Or i = 4 Or i = 5 Or i = 6 Then
        shpWuerfelAuge(1).Visible = True
        shpWuerfelAuge(2).Visible = True
    Else
        shpWuerfelAuge(1).Visible = False
        shpWuerfelAuge(2).Visible = False
    End If
    If i = 3 Or i = 4 Or i = 5 Or i = 6 Then
        shpWuerfelAuge(3).Visible = True
        shpWuerfelAuge(4).Visible = True
    Else
        shpWuerfelAuge(3).Visible = False
        shpWuerfelAuge(4).Visible = False
    End If
    If i = 6 Then
        shpWuerfelAuge(5).Visible = True
        shpWuerfelAuge(6).Visible = True
    Else
        shpWuerfelAuge(5).Visible = False
        shpWuerfelAuge(6).Visible = False
    End If
End Sub

Private Sub VBSleep(ByVal dwMilliseconds As Long)
    Dim StopTime As Single
    StopTime = Timer + dwMilliseconds / 1000
    Do While Timer < StopTime
        Call Sleep(10)
        DoEvents
    Loop
End Sub

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    Call StartGame
End Sub

Private Sub Form_Click()
'    If NextStep = False Then
'        Timer1.Enabled = False
'        ZeigeHinweis "Das Spiel ist beendet"
'        cmdStart.Enabled = True
'    End If
End Sub

Private Sub Form_Load()
    Call InitializeGame
    ZeigeHinweis "Zum Start den Startbutton drücken..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Call VBSleep(500)
End Sub

Private Sub Timer1_Timer()
    If NextStep = False Then
        Timer1.Enabled = False
        'MsgBox "Das Spiel ist beendet"
        ZeigeHinweis "Das Spiel ist beendet"
        cmdStart.Enabled = True
        'Call InitializeGame
        'Call StartGame
    End If
End Sub
