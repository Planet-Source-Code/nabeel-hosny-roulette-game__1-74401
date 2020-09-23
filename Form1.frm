VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008080&
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14940
   ControlBox      =   0   'False
   DrawWidth       =   2
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   760
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   996
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox HelpMe 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2825
      Left            =   4590
      MultiLine       =   -1  'True
      TabIndex        =   61
      Text            =   "Form1.frx":246042
      Top             =   4050
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox listwin 
      Height          =   255
      ItemData        =   "Form1.frx":2461D2
      Left            =   9600
      List            =   "Form1.frx":246245
      TabIndex        =   58
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox list 
      Height          =   2400
      Index           =   1
      ItemData        =   "Form1.frx":246405
      Left            =   0
      List            =   "Form1.frx":24642D
      TabIndex        =   57
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   2400
      Index           =   2
      ItemData        =   "Form1.frx":246461
      Left            =   360
      List            =   "Form1.frx":246489
      TabIndex        =   56
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   2400
      Index           =   3
      ItemData        =   "Form1.frx":2464BD
      Left            =   720
      List            =   "Form1.frx":2464E5
      TabIndex        =   55
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   3570
      Index           =   4
      ItemData        =   "Form1.frx":246519
      Left            =   1200
      List            =   "Form1.frx":246553
      TabIndex        =   54
      Tag             =   "2"
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   3570
      Index           =   5
      ItemData        =   "Form1.frx":24659F
      Left            =   1560
      List            =   "Form1.frx":2465D9
      TabIndex        =   53
      Tag             =   "2"
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   2400
      Index           =   6
      ItemData        =   "Form1.frx":246625
      Left            =   2040
      List            =   "Form1.frx":24664D
      TabIndex        =   52
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   2400
      Index           =   9
      ItemData        =   "Form1.frx":246681
      Left            =   3240
      List            =   "Form1.frx":2466A9
      TabIndex        =   51
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   3570
      Index           =   10
      ItemData        =   "Form1.frx":2466DD
      Left            =   3840
      List            =   "Form1.frx":246717
      TabIndex        =   50
      Tag             =   "2"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   3570
      Index           =   11
      ItemData        =   "Form1.frx":246763
      Left            =   4200
      List            =   "Form1.frx":24679D
      TabIndex        =   49
      Tag             =   "2"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   2400
      Index           =   12
      ItemData        =   "Form1.frx":2467E9
      Left            =   4560
      List            =   "Form1.frx":246811
      TabIndex        =   48
      Tag             =   "3"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   1425
      Index           =   8
      ItemData        =   "Form1.frx":246845
      Left            =   2880
      List            =   "Form1.frx":24685E
      TabIndex        =   47
      Tag             =   "9"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox list 
      Height          =   1815
      Index           =   7
      ItemData        =   "Form1.frx":24687E
      Left            =   2400
      List            =   "Form1.frx":24689D
      TabIndex        =   46
      Tag             =   "7"
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox listupdate 
      Height          =   255
      Left            =   8400
      TabIndex        =   39
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808000&
      Caption         =   "HELP "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   840
      MouseIcon       =   "Form1.frx":2468C5
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   9720
      Width           =   1200
   End
   Begin VB.PictureBox startme 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   10920
      MouseIcon       =   "Form1.frx":246BCF
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":246D21
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   35
      Top             =   10440
      Width           =   1545
   End
   Begin VB.PictureBox exitme 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   13200
      MouseIcon       =   "Form1.frx":24B413
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":24B565
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   33
      Top             =   10440
      Width           =   1545
   End
   Begin VB.Frame fra 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7920
      TabIndex        =   14
      Top             =   9600
      Width           =   2730
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   4
         FillColor       =   &H0000FFFF&
         Height          =   255
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label Win 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   45
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Win 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   44
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Win 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   43
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Win 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   42
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Win 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   41
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   40
         Top             =   120
         Width           =   600
      End
      Begin VB.Label LuckNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1300
         TabIndex        =   32
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label LuckNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1300
         TabIndex        =   31
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label LuckNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1300
         TabIndex        =   30
         Top             =   840
         Width           =   750
      End
      Begin VB.Label LuckNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1300
         TabIndex        =   29
         Top             =   600
         Width           =   750
      End
      Begin VB.Label LuckNo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1300
         TabIndex        =   28
         Top             =   360
         Width           =   750
      End
      Begin VB.Label BitMon 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 $"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   572
         TabIndex        =   27
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label BitMon 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 $"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   572
         TabIndex        =   26
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label BitMon 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 $"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   572
         TabIndex        =   25
         Top             =   840
         Width           =   750
      End
      Begin VB.Label BitMon 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 $"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   572
         TabIndex        =   24
         Top             =   600
         Width           =   750
      End
      Begin VB.Label BitMon 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "000 $"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   572
         TabIndex        =   23
         Top             =   360
         Width           =   750
      End
      Begin VB.Label bitno 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   80
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label bitno 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   80
         TabIndex        =   21
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label bitno 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   80
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.Label bitno 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   80
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label bitno 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Luck No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1300
         TabIndex        =   17
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bit Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   572
         TabIndex        =   16
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bit No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   80
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Bit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   4
      Left            =   1800
      MouseIcon       =   "Form1.frx":24FC57
      MousePointer    =   99  'Custom
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   11
      Top             =   8520
      Width           =   675
   End
   Begin VB.PictureBox Bit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   3
      Left            =   840
      MouseIcon       =   "Form1.frx":24FDA9
      MousePointer    =   99  'Custom
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   10
      Top             =   9300
      Width           =   675
   End
   Begin VB.PictureBox Bit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   2
      Left            =   840
      MouseIcon       =   "Form1.frx":24FEFB
      MousePointer    =   99  'Custom
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   9
      Top             =   9300
      Width           =   675
   End
   Begin VB.PictureBox Bit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   1
      Left            =   840
      MouseIcon       =   "Form1.frx":25004D
      MousePointer    =   99  'Custom
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   8
      Top             =   9300
      Width           =   675
   End
   Begin VB.PictureBox Bit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   0
      Left            =   840
      MouseIcon       =   "Form1.frx":25019F
      MousePointer    =   99  'Custom
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   7
      Top             =   9300
      Width           =   675
   End
   Begin VB.ListBox listbit 
      Height          =   1035
      ItemData        =   "Form1.frx":2502F1
      Left            =   120
      List            =   "Form1.frx":250304
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "Form1.frx":250317
      Left            =   9600
      List            =   "Form1.frx":25038A
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   255
      ItemData        =   "Form1.frx":2504DB
      Left            =   9600
      List            =   "Form1.frx":25073A
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   11280
      TabIndex        =   2
      Top             =   10680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Slipinng"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   1
      Top             =   10680
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "Form1.frx":250E4F
      Left            =   8400
      List            =   "Form1.frx":250EC2
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSBar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00CA8826&
      BackStyle       =   0  'Transparent
      Caption         =   "nabeelhosny@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   60
      Top             =   11160
      Width           =   2730
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Bits 5 Times in 1 Round"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   3240
      TabIndex        =   59
      Top             =   120
      Width           =   5235
   End
   Begin VB.Shape Shwin 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Height          =   750
      Left            =   12795
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   2
      Left            =   13770
      Picture         =   "Form1.frx":251082
      Stretch         =   -1  'True
      Top             =   9690
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   1
      Left            =   13035
      Picture         =   "Form1.frx":25114E
      Stretch         =   -1  'True
      Top             =   9690
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   0
      Left            =   12240
      Picture         =   "Form1.frx":25121A
      Stretch         =   -1  'True
      Top             =   9690
      Width           =   345
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   360
      Left            =   13110
      TabIndex        =   37
      Top             =   120
      Width           =   105
   End
   Begin VB.Label Labcash 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00500"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Left            =   9420
      TabIndex        =   36
      Top             =   1230
      Width           =   1140
   End
   Begin VB.Label Cashb4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0000"
      Height          =   255
      Left            =   14400
      TabIndex        =   34
      Top             =   240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      Height          =   255
      Left            =   10800
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00 00"
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image money 
      Height          =   675
      Index           =   4
      Left            =   600
      MouseIcon       =   "Form1.frx":2512E6
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":251438
      Tag             =   "100"
      Top             =   10500
      Width           =   675
   End
   Begin VB.Image money 
      Height          =   675
      Index           =   3
      Left            =   1440
      MouseIcon       =   "Form1.frx":252C62
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":252DB4
      Tag             =   "025"
      Top             =   10500
      Width           =   675
   End
   Begin VB.Image money 
      Height          =   675
      Index           =   2
      Left            =   2280
      MouseIcon       =   "Form1.frx":2545DE
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":254730
      Tag             =   "010"
      Top             =   10500
      Width           =   675
   End
   Begin VB.Image money 
      Height          =   675
      Index           =   1
      Left            =   3120
      MouseIcon       =   "Form1.frx":255F5A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":2560AC
      Tag             =   "005"
      Top             =   10500
      Width           =   675
   End
   Begin VB.Image money 
      Height          =   675
      Index           =   0
      Left            =   3960
      MouseIcon       =   "Form1.frx":2578D6
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":257A28
      Tag             =   "001"
      Top             =   10500
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   840
      Left            =   1245
      TabIndex        =   4
      Top             =   1080
      Width           =   390
   End
   Begin VB.Image dalil 
      Height          =   480
      Index           =   0
      Left            =   7800
      Picture         =   "Form1.frx":259252
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   3120
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image centre 
      Height          =   480
      Left            =   5490
      Picture         =   "Form1.frx":25955C
      Stretch         =   -1  'True
      Top             =   5220
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI = 3.14159
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Dim mblnRunning As Boolean  'Is the render loop running?
Dim ddd, bb, ik As Integer
Dim a(5) As String


Private Sub Bit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If mblnRunning = True Then Exit Sub
Xmouse = X: oldleft = Bit(Index).Left
Ymouse = Y: oldtop = Bit(Index).Top
listbit.ListIndex = Bit(Index).Index
End Sub

Private Sub Bit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If mblnRunning = True Then Exit Sub
If Button = 1 Then
x1 = Bit(Index).Left + X - Xmouse
y1 = Bit(Index).Top + Y - Ymouse
Bit(Index).Move x1 - 10, y1 - 10
Bit(Index).ZOrder (0)
End If
tx = Int((Bit(Index).Left - 804) / 48) + 1
ty = Int((Bit(Index).Top - 90) / 46) + 1
If tx > 0 And tx < 4 And ty < 14 And ty > 0 Then
Label2.Caption = Format(tx - 1, "00") & " " & Format(ty - 1, "00")
Label3.Caption = Format(tx + (ty - 1) * 3, "00")
Label7.Caption = "NumR"
End If

If tx = 1 And ty = 13 Then Label7.Caption = "Col1"
If tx = 2 And ty = 13 Then Label7.Caption = "Col2"
If tx = 3 And ty = 13 Then Label7.Caption = "Col3"

If tx > 0 And tx < 4 And ty = -1 Then Label7.Caption = "Num0": Label3.Caption = "00"

If Bit(Index).Left + 20 > 722 And Bit(Index).Left + 20 < 763 And Bit(Index).Top + 20 > 88 And Bit(Index).Top + 20 < 180 Then Label7.Caption = "01:18"
If Bit(Index).Left + 20 > 722 And Bit(Index).Left + 20 < 763 And Bit(Index).Top + 20 > 180 And Bit(Index).Top + 20 < 248 Then Label7.Caption = "Even"
If Bit(Index).Left + 20 > 763 And Bit(Index).Left + 20 < 804 And Bit(Index).Top + 20 > 88 And Bit(Index).Top + 20 < 248 Then Label7.Caption = "01:12"

If Bit(Index).Left + 20 > 722 And Bit(Index).Left + 20 < 763 And Bit(Index).Top + 20 > 248 And Bit(Index).Top + 20 < 366 Then Label7.Caption = "Quad"
If Bit(Index).Left + 20 > 722 And Bit(Index).Left + 20 < 763 And Bit(Index).Top + 20 > 366 And Bit(Index).Top + 20 < 460 Then Label7.Caption = "Pent"
If Bit(Index).Left + 20 > 763 And Bit(Index).Left + 20 < 804 And Bit(Index).Top + 20 > 248 And Bit(Index).Top + 20 < 460 Then Label7.Caption = "13:24"

If Bit(Index).Left + 20 > 722 And Bit(Index).Left + 20 < 763 And Bit(Index).Top + 20 > 460 And Bit(Index).Top + 20 < 552 Then Label7.Caption = "ODD"
If Bit(Index).Left + 20 > 722 And Bit(Index).Left + 20 < 763 And Bit(Index).Top + 20 > 552 And Bit(Index).Top + 20 < 644 Then Label7.Caption = "19:36"
If Bit(Index).Left + 20 > 763 And Bit(Index).Left + 20 < 804 And Bit(Index).Top + 20 > 460 And Bit(Index).Top + 20 < 644 Then Label7.Caption = "25:36"

If Bit(Index).Left < 722 Or Bit(Index).Left > 954 Or _
   Bit(Index).Left < 804 And Bit(Index).Top < 90 Or _
   Bit(Index).Left < 804 And Bit(Index).Top > 644 Then _
   Label2.Caption = "": Label3.Caption = "": Label7.Caption = ""

End Sub

Private Sub Bit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label7.Caption = "" Then Bit(Index).Move 56, 620: Bit(Index).Visible = False: Exit Sub
Select Case Label7.Caption
Case "NumR": Bit(Index).Move 805 + Left(Label2.Caption, 2) * 50, 90 + Right(Label2.Caption, 2) * 46: Bit(Index).Tag = Label3.Caption: a(Bit(Index).Index) = 0
Case "Col1": Bit(Index).Move 805 + Left(Label2.Caption, 2) * 50, 90 + Right(Label2.Caption, 2) * 46: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 1
Case "Col2": Bit(Index).Move 805 + Left(Label2.Caption, 2) * 50, 90 + Right(Label2.Caption, 2) * 46: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 2
Case "Col3": Bit(Index).Move 805 + Left(Label2.Caption, 2) * 50, 90 + Right(Label2.Caption, 2) * 46: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 3
Case "Num0": Bit(Index).Move 855, 44: Bit(Index).Tag = "00": a(Bit(Index).Index) = 0
Case "01:18": Bit(Index).Move 720, 112: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 4
Case "Even": Bit(Index).Move 720, 200: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 5
Case "01:12": Bit(Index).Move 761, 228: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 6
Case "Quad": Bit(Index).Move 720, 298: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 7
Case "Pent": Bit(Index).Move 720, 390: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 8
Case "13:24": Bit(Index).Move 761, 414: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 9
Case "ODD": Bit(Index).Move 720, 484: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 10
Case "19:36": Bit(Index).Move 720, 576: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 11
Case "25:36": Bit(Index).Move 761, 600: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = 12
Case eles: Bit(Index).Move 56, 620: Bit(inddex).Visible = False: Bit(Index).Tag = Label7.Caption: a(Bit(Index).Index) = ""
End Select

If listbit.ListIndex <> 4 Then listbit.ListIndex = listbit.ListIndex + 1
End Sub

Private Sub exitme_Click()
mblnRunning = False
End
End Sub

Private Sub Form_Load()
 Me.WindowState = 2 'final1
 Me.Move 0, 0, Screen.Width, Screen.Height
 For i = 4 To 0 Step -1: Bit(i).Move 56, 620: Bit(i).ZOrder: Bit(i).Visible = False: Next
 For i = 0 To 4: SetWindowRgn Bit(i).hwnd, CreateRoundRectRgn(0, 0, Bit(i).Width, Bit(i).Height, 50, 50), True: Next
SetWindowRgn fra.hwnd, CreateRoundRectRgn(0, 0, fra.Width, fra.Height, 50, 50), True
SetWindowRgn exitme.hwnd, CreateRoundRectRgn(0, 0, exitme.Width, exitme.Height, 30, 30), True
SetWindowRgn startme.hwnd, CreateRoundRectRgn(0, 0, startme.Width, startme.Height, 30, 30), True
SetWindowRgn HelpMe.hwnd, CreateRoundRectRgn(0, 0, HelpMe.Width, HelpMe.Height, 30, 30), True
SetWindowRgn Check1.hwnd, CreateRoundRectRgn(0, 0, Check1.Width, Check1.Height, 10, 10), True
Me.BackColor = RGB(192, 192, 192)
For i = 1 To 36: Load img(i): img(i).Visible = True: Next
 For i = 0 To List1.ListCount - 1
 img(Left(List1.list(i), 2)).Move Mid(List1.list(i), 4, 3), Right(List1.list(i), 3)
 img(Left(List1.list(i), 2)).Picture = LoadPicture(App.Path & "\pic\" & Format(Left(List1.list(i), 2), "00") & ".ico")
 img(Left(List1.list(i), 2)).ZOrder (0)
 Next
centre.Move 364, 348
Shape2.Visible = False
Cashb4.Caption = Labcash.Caption
Shwin.Visible = False
    mblnRunning = True
    Me.Show
    ddd = 0
listbit.ListIndex = 0
Option1(1).Value = True

End Sub
Sub run()
Shape2.Visible = False

   Randomize
    List2.Selected(Int(List2.ListCount * Rnd)) = True
  DoEvents
 pal = PI / 100
Do While mblnRunning

             drawship                    'Draw the ship
            dodo
            Sleep ddd * 0.025
        'Allow other events to occur
        
        If ddd >= 400 + Int(200 * Rnd) Then mblnRunning = False
        ddd = ddd + 1
      DoEvents
  sndPlaySound App.Path & "\pic\in.wav", 1: DoEvents
    Loop
    If mblnRunning = False Then checkme
End Sub
Sub dodo()
pal = PI / 100

Distance1 = Sqr(((dalil(0).Left - centre.Left) ^ 2) + ((centre.Top - dalil(0).Top) ^ 2))
dx = (dalil(0).Left - centre.Left)
dy = (dalil(0).Top - centre.Top)
        If dx = 0 Then dx = 0.00001
        If dy = 0 Then dy = 0.00001
xdeg = Atn(dy / dx)
        If dx > 0 And dy < 0 Then adeg = 6.283 - (xdeg * -1) '+,-
        If dx < 0 And dy < 0 Then adeg = 3.142 + (xdeg * 1) '-,-
        If dx < 0 And dy > 0 Then adeg = 3.142 - (xdeg * -1) '-,+
        If dx > 0 And dy > 0 Then adeg = (xdeg * 1)          '+,+
  dalil(0).Move Distance1 * Cos(adeg - pal) + centre.Left, centre.Top + Distance1 * Sin(adeg - pal)
DoEvents

End Sub
Sub drawship()
pal = PI / 100
 
For i = 0 To img.UBound
Distance = Sqr(((img(i).Left - centre.Left) ^ 2) + ((centre.Top - img(i).Top) ^ 2))

dx1 = (img(i).Left - centre.Left)
dy1 = (img(i).Top - centre.Top)
        If dx1 = 0 Then dx1 = 0.00001
        If dy1 = 0 Then dy1 = 0.00001
xdeg1 = Atn(dy1 / dx1)
        If dx1 > 0 And dy1 < 0 Then adeg1 = 6.283 - (xdeg1 * -1) '+,-
        If dx1 < 0 And dy1 < 0 Then adeg1 = 3.142 + (xdeg1 * 1)  '-,-
        If dx1 < 0 And dy1 > 0 Then adeg1 = 3.142 - (xdeg1 * -1) '-,+
        If dx1 > 0 And dy1 > 0 Then adeg1 = (xdeg1 * 1)          '+,+
img(i).Move Distance * Cos(adeg1 + pal) + centre.Left, centre.Top + Distance * Sin(adeg1 + pal)
DoEvents
 Next

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = ""
End Sub

Private Sub List2_Click()
dalil(0).Move Val(Left(List2.list(List2.ListIndex), 3)), Val(Right(List2.list(List2.ListIndex), 3))
End Sub

Private Sub List3_Click()
dalil(0).Move Val(Left(List3.list(List3.ListIndex), 3)), Val(Right(List3.list(List3.ListIndex), 3))
End Sub

Private Sub LuckNo_Click(Index As Integer)
Label8.Caption = LuckNo(Index).Tag
End Sub

Private Sub money_Click(Index As Integer)
If mblnRunning = True Then Exit Sub
Bit(listbit.list(listbit.ListIndex)).Visible = True
Bit(listbit.list(listbit.ListIndex)).Move money(Index).Left, money(Index).Top - 20
 Bit(listbit.list(listbit.ListIndex)).Picture = money(Index).Picture
 Bit(listbit.list(listbit.ListIndex)).ToolTipText = money(Index).Tag & " $"
End Sub



Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
mblnRunning = True
ddd = 0
run
Else
mblnRunning = False
End If
End Sub


Sub checkme()
Dim kk As Integer
updateme
DoEvents

Label1.Caption = ""
mblnRunning = False
For i = 0 To List3.ListCount - 1
  If Abs(Int(dalil(0).Left) - Val(Left(List3.list(i), 3))) < 12 And _
     Abs(Int(dalil(0).Top) - Val(Right(List3.list(i), 3))) < 12 Then _
  List3.Selected(i) = True: Exit For
Next

If List3.SelCount = 0 Then List3.ListIndex = 0

For ik = 0 To 36
listupdate.Selected(ik) = True
kk = Val(Left(listupdate.list(listupdate.ListIndex), 2))
Distance2 = Sqr(((img(kk).Left - dalil(0).Left) ^ 2) + ((dalil(0).Top - img(kk).Top) ^ 2))
If Val(Distance2) < 32 Then Label1.Caption = Format(kk, "00"): Exit For
DoEvents
Next
'
Shwin.Visible = True
For ik = 0 To 36
If Val(Label1.Caption) = Val(Left(listwin.list(ik), 3)) Then _
Shwin.Move Val(Mid(listwin.list(ik), 4, 3)), Val(Right(listwin.list(ik), 3)): Exit For
Next
'
Option1(1).Value = True


For i = 0 To listbit.ListIndex
Select Case LuckNo(i).ToolTipText
Case "0"
If Val(LuckNo(i).Caption) = Val(Label1.Caption) Then _
   Labcash.Caption = Format(Val(Labcash.Caption) + Val(Left(BitMon(i).Caption, 3)) * 36, "00000"): _
   Shape2.Visible = True: Shape2.Top = bitno(i).Top + 30: _
   sndPlaySound App.Path & "\pic\yes.wav", 1: DoEvents
Case ""
 sndPlaySound App.Path & "\pic\wrong.wav", 1: DoEvents
Case Else
For ii = 0 To list(Val(LuckNo(i).ToolTipText)).ListCount - 1
list(Val(LuckNo(i).ToolTipText)).Selected(ii) = True
 If Val(Label1.Caption) = Val(list(Val(LuckNo(i).ToolTipText)).list(ii)) Then _
 Labcash.Caption = Format(Val(Labcash.Caption) + (Val(Left(BitMon(i).Caption, 3)) * Val(list(Val(LuckNo(i).ToolTipText)).Tag)), "00000"): _
 Shape2.Visible = True: Shape2.Top = bitno(i).Top + 30: _
 sndPlaySound App.Path & "\pic\yes.wav", 1: DoEvents
 DoEvents
Next ii
End Select
DoEvents
Next
clearme

End Sub

Sub clearme()

If Val(Cashb4.Caption) > Val(Labcash.Caption) Then
MsgBox "Rulette has been Stop !" & vbCrLf & _
       "    On Number " & Label1.Caption & vbCrLf & _
       "  Bank Win  " & Format(Val(Cashb4.Caption) - Val(Labcash.Caption), "0000") & " $", vbOKOnly
Else
MsgBox "Rulette has been Stop !" & vbCrLf & _
       "    On Number " & Label1.Caption & vbCrLf & _
       "  player Win  " & Format(Val(Labcash.Caption) - Val(Cashb4.Caption), "0000") & " $", vbOKOnly
End If
Shwin.Visible = False
 'Sleep 5000
For i = 0 To 4
Bit(i).Picture = Nothing
Bit(i).Visible = False
Bit(i).ToolTipText = ""
Bit(i).Tag = ""
Win(i).Caption = "000"
bitno(i).Caption = "00"
BitMon(i).Caption = "000 $"
LuckNo(i).Caption = "00"
DoEvents
Next
Shape2.Visible = False
For i = 4 To 0 Step -1: Bit(i).Move 56, 620: Bit(i).ZOrder: Next
listbit.ListIndex = 0: Label1.Caption = "  "
End Sub
'
Private Sub startme_Click()
Dim num As Integer
num = 0 ': Shwin.Visible = False
Cashb4.Caption = Labcash.Caption
For i = 0 To 4
If Bit(i).Left > 700 Then Labcash.Caption = Format(Val(Labcash.Caption) - Val(Left(Bit(i).ToolTipText, 3)), "00000"): num = num + 1
  Win(i).Caption = "0000"
  bitno(i).Caption = "00"
  BitMon(i).Caption = "000 $"
  LuckNo(i).Caption = "00"
 ' a(i) = ""
DoEvents
Next

listbit.ListIndex = num
  For i = 0 To listbit.ListIndex - 1
  bitno(i).Caption = Format(i + 1, "00")
  BitMon(i).Caption = Bit(i).ToolTipText
  LuckNo(i).Caption = Bit(i).Tag
  LuckNo(i).ToolTipText = a(i)
  If a(i) <> 0 Then Win(i).Caption = Format((Val(list(a(i)).Tag) - 1) * Val(Left(BitMon(i).Caption, 3)), "0000")
  If a(i) = 0 Then Win(i).Caption = Format(35 * Val(Left(BitMon(i).Caption, 3)), "0000")
   Next

If listbit.ListIndex = 0 Then
MsgBox "Put Your Bit First !" & vbCrLf, vbOKOnly
 Exit Sub
 End If
Option1(0).Value = True
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
HelpMe.Visible = True: HelpMe.ZOrder (0)
Else
HelpMe.Visible = False
End If
End Sub
Sub updateme()
Dim x1, y1 As Long
Dim x2, y2, i, ii, cc As Integer
listupdate.Clear

x1 = img(0).Left: y1 = img(0).Top
  For i = 0 To List1.ListCount - 1
  x2 = Val(Mid(List1.list(i), 4, 3)): y2 = Val(Right(List1.list(i), 3))
   If Int(Abs(x1 - x2)) < 16 And Int(Abs(y1 - y2)) < 16 Then _
     img(0).Move x2, y2: List1.ListIndex = i: Exit For
    DoEvents
Next
cc = -1
For i = List1.ListIndex To List1.ListCount - 1
cc = cc + 1
'img(Left(List1.List(i), 2)).Move Val(Mid(List1.List(i), 4, 3)), Val(Right(List1.List(i), 3))
listupdate.AddItem Format(Left(List1.list(cc), 2), "00") & " " & Format(Mid(List1.list(i), 4, 3), "000") & " " & Format(Right(List1.list(i), 3), "000")

DoEvents
Next
For i = 0 To List1.ListIndex - 1
cc = cc + 1
'img(Left(List1.List(i), 2)).Move Val(Mid(List1.List(i), 4, 3)), Val(Right(List1.List(i), 3))
listupdate.AddItem Format(Left(List1.list(cc), 2), "00") & " " & Format(Mid(List1.list(i), 4, 3), "000") & " " & Format(Right(List1.list(i), 3), "000")
DoEvents
Next

 For i = 0 To listupdate.ListCount - 1
 img(Left(listupdate.list(i), 2)).Move Val(Mid(listupdate.list(i), 4, 3)), Val(Right(listupdate.list(i), 3))
 DoEvents
 Next
baba:
'DoEvents
'Next

End Sub


