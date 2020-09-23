VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00494949&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New MenuColors thru subclassing"
   ClientHeight    =   5130
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By this way i could change the Position of the Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   360
      TabIndex        =   10
      Top             =   3960
      Width           =   3660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu is marked as PopUp without any entry"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   3120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move menu to right :-)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   1440
      TabIndex        =   8
      Top             =   3480
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click File / New to call an original MenuRoutine"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right Click on form for PopUp"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PopUp Menus are possible"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use the original ...Click Sub"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uses Standard Menus"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text ,Back, Selction, Disabled"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Colors for the menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1965
   End
   Begin VB.Menu MnuPop23 
      Caption         =   "Move menu to right :-)"
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuFileOpen 
         Caption         =   "Open"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileSave 
         Caption         =   "Save"
         Begin VB.Menu MnuFileNormal 
            Caption         =   "Normal"
         End
         Begin VB.Menu MnuFileSaveAs 
            Caption         =   "Save As"
         End
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuHowTo 
         Caption         =   "How to"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu MnuPopUp01 
      Caption         =   "PopUp01"
      Begin VB.Menu MnuDum01 
         Caption         =   "Dummy01"
      End
      Begin VB.Menu MnuDum02 
         Caption         =   "Dummy02"
      End
      Begin VB.Menu MnuDum03 
         Caption         =   "Dummy03"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'See the Module for more Infos
Private Sub Form_Load()

'Call th Subclassing routines

    ChangeMenu Me, &HFFFFFF, &H494949, &HFF0000, &HFFFFFF

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'Standard PopUp

    If Button = 2 Then
        PopupMenu MnuPopUp01
    End If

End Sub

Private Sub MnuFileNew_Click()

    MsgBox "New File selected :-)", vbInformation, "Color Menu Example"

End Sub
