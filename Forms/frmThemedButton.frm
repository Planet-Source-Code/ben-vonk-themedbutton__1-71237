VERSION 5.00
Begin VB.Form frmThemedButton 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demo ThemedButton"
   ClientHeight    =   5892
   ClientLeft      =   36
   ClientTop       =   516
   ClientWidth     =   7332
   Icon            =   "frmThemedButton.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   2  'CenterScreen
   Begin prjThemedButton.ThemedButton thbDemo 
      Height          =   420
      Left            =   3000
      TabIndex        =   23
      Top             =   5400
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   741
      ButtonCorner    =   1
      ButtonDefaulted =   "frmThemedButton.frx":058A
      ButtonNormal    =   "frmThemedButton.frx":2C88
      ButtonOver      =   "frmThemedButton.frx":55DE
      ButtonPressed   =   "frmThemedButton.frx":7BB0
      ButtonRounding  =   9
      ButtonThemeType =   1
      Caption         =   "&Demo"
      CaptionShadow   =   -1  'True
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   14737632
      MouseIcon       =   "frmThemedButton.frx":A182
      Value           =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   252
      Left            =   5040
      TabIndex        =   22
      Top             =   4920
      Value           =   2  'Grayed
      Width           =   1812
   End
   Begin prjThemedButton.ThemedButton ThemedButton16 
      Height          =   216
      Left            =   5040
      TabIndex        =   21
      Top             =   4560
      Width           =   1812
      _ExtentX        =   3196
      _ExtentY        =   381
      ButtonDisabled  =   "frmThemedButton.frx":A71C
      ButtonDisabledGrayed=   "frmThemedButton.frx":B36E
      ButtonDisabledValued=   "frmThemedButton.frx":BFC0
      ButtonNormal    =   "frmThemedButton.frx":CC12
      ButtonNormalGrayed=   "frmThemedButton.frx":D864
      ButtonNormalValued=   "frmThemedButton.frx":E4B6
      ButtonOver      =   "frmThemedButton.frx":F108
      ButtonOverGrayed=   "frmThemedButton.frx":FD5A
      ButtonOverValued=   "frmThemedButton.frx":109AC
      ButtonPressed   =   "frmThemedButton.frx":115FE
      ButtonPressedGrayed=   "frmThemedButton.frx":12250
      ButtonPressedValued=   "frmThemedButton.frx":12EA2
      ButtonThemeType =   1
      ButtonType      =   2
      CaptionAlign    =   0
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":13AF4
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
   End
   Begin prjThemedButton.ThemedButton ThemedButton15 
      Height          =   252
      Left            =   5040
      TabIndex        =   20
      Top             =   4200
      Width           =   1812
      _ExtentX        =   3196
      _ExtentY        =   445
      ButtonType      =   2
      CaptionAlign    =   1
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":1408E
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
   End
   Begin prjThemedButton.ThemedButton ThemedButton14 
      Height          =   252
      Left            =   5040
      TabIndex        =   19
      Top             =   3840
      Width           =   1932
      _ExtentX        =   3408
      _ExtentY        =   445
      ButtonType      =   2
      CaptionAlign    =   0
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":14628
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   192
      Left            =   5040
      TabIndex        =   18
      Top             =   2640
      Width           =   1932
   End
   Begin prjThemedButton.ThemedButton ThemedButton13 
      Height          =   252
      Index           =   0
      Left            =   5040
      TabIndex        =   17
      Top             =   2160
      Width           =   1932
      _ExtentX        =   3408
      _ExtentY        =   445
      ButtonDisabled  =   "frmThemedButton.frx":14BC2
      ButtonDisabledValued=   "frmThemedButton.frx":15814
      ButtonNormal    =   "frmThemedButton.frx":16466
      ButtonNormalValued=   "frmThemedButton.frx":170B8
      ButtonOver      =   "frmThemedButton.frx":17D0A
      ButtonOverValued=   "frmThemedButton.frx":1895C
      ButtonPressed   =   "frmThemedButton.frx":195AE
      ButtonPressedValued=   "frmThemedButton.frx":1A200
      ButtonThemeType =   1
      ButtonType      =   1
      Caption         =   "ThemedButton13"
      CaptionAlign    =   1
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":1AE52
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
      Value           =   0
   End
   Begin prjThemedButton.ThemedButton ThemedButton12 
      Height          =   252
      Left            =   5040
      TabIndex        =   16
      Top             =   1800
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      ButtonType      =   1
      CaptionAlign    =   0
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":1B3EC
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
      Value           =   0
   End
   Begin prjThemedButton.ThemedButton ThemedButton10 
      Height          =   252
      Index           =   0
      Left            =   5040
      TabIndex        =   12
      Top             =   240
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      ButtonType      =   1
      Caption         =   "ThemedButton10 Array"
      CaptionAlign    =   0
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":1B986
      OptionButtonMultiSelect=   -1  'True
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
      Value           =   0
   End
   Begin prjThemedButton.ThemedButton ThemedButton9 
      Height          =   216
      Left            =   1680
      TabIndex        =   1
      Top             =   5520
      Width           =   516
      _ExtentX        =   910
      _ExtentY        =   381
      ButtonDisabled  =   "frmThemedButton.frx":1BF20
      ButtonNormal    =   "frmThemedButton.frx":1C8BA
      ButtonOver      =   "frmThemedButton.frx":1D254
      ButtonPressed   =   "frmThemedButton.frx":1DBEE
      ButtonRounding  =   6
      ButtonThemeType =   1
      Caption         =   ""
      Enabled         =   0   'False
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":1E588
   End
   Begin prjThemedButton.ThemedButton thbQuit 
      Height          =   216
      Left            =   360
      TabIndex        =   0
      Top             =   5520
      Width           =   516
      _ExtentX        =   910
      _ExtentY        =   381
      ButtonDisabled  =   "frmThemedButton.frx":1EB22
      ButtonNormal    =   "frmThemedButton.frx":1F4BC
      ButtonOver      =   "frmThemedButton.frx":1FE56
      ButtonPressed   =   "frmThemedButton.frx":207F0
      ButtonRounding  =   6
      ButtonThemeType =   1
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":2118A
      ShowFocusRect   =   0   'False
   End
   Begin prjThemedButton.ThemedButton ThemedButton8 
      Height          =   612
      Left            =   2640
      TabIndex        =   11
      Top             =   4560
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1080
      CaptionShadow   =   -1  'True
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   32768
      MouseIcon       =   "frmThemedButton.frx":21724
   End
   Begin prjThemedButton.ThemedButton ThemedButton7 
      Height          =   612
      Left            =   2640
      TabIndex        =   10
      Top             =   3840
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1080
      CaptionAlign    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      MouseIcon       =   "frmThemedButton.frx":21CBE
      Picture         =   "frmThemedButton.frx":22258
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin prjThemedButton.ThemedButton ThemedButton6 
      Height          =   432
      Left            =   2640
      TabIndex        =   9
      Top             =   3120
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   339
      CaptionAlign    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":227F2
   End
   Begin prjThemedButton.ThemedButton ThemedButton5 
      Height          =   612
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1080
      CaptionAlign    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":22D8C
      Picture         =   "frmThemedButton.frx":23326
      PictureAlign    =   3
   End
   Begin prjThemedButton.ThemedButton ThemedButton4 
      Height          =   852
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1503
      CaptionShadow   =   -1  'True
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":23C00
      Picture         =   "frmThemedButton.frx":2419A
      PictureAlign    =   4
   End
   Begin prjThemedButton.ThemedButton ThemedButton3 
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   656
      CaptionAlign    =   0
      CaptionShadow   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":24A74
   End
   Begin prjThemedButton.ThemedButton ThemedButton2 
      Height          =   1092
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1926
      Caption         =   "Themed&Button2"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":2500E
      Picture         =   "frmThemedButton.frx":255A8
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Com&mand2"
      Enabled         =   0   'False
      Height          =   1092
      Left            =   2640
      Picture         =   "frmThemedButton.frx":25E82
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   2640
      Picture         =   "frmThemedButton.frx":2674C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   2052
   End
   Begin prjThemedButton.ThemedButton ThemedButton1 
      Height          =   1212
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   2138
      Caption         =   "&ThemedButton1"
      CaptionShadow   =   -1  'True
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16744576
      Picture         =   "frmThemedButton.frx":27016
      ShowFocusRect   =   0   'False
   End
   Begin prjThemedButton.ThemedButton ThemedButton10 
      Height          =   252
      Index           =   1
      Left            =   5040
      TabIndex        =   13
      Top             =   600
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      ButtonType      =   1
      Caption         =   "ThemedButton10 Array"
      CaptionAlign    =   0
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":278F0
      OptionButtonMultiSelect=   -1  'True
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
      Value           =   0
   End
   Begin prjThemedButton.ThemedButton ThemedButton10 
      Height          =   252
      Index           =   2
      Left            =   5040
      TabIndex        =   14
      Top             =   960
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      ButtonType      =   1
      Caption         =   "ThemedButton10 Array"
      CaptionAlign    =   0
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":27E8A
      OptionButtonMultiSelect=   -1  'True
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
      Value           =   0
   End
   Begin prjThemedButton.ThemedButton ThemedButton11 
      Height          =   252
      Left            =   5040
      TabIndex        =   15
      Top             =   1320
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   445
      ButtonType      =   1
      Caption         =   "ThemedButton11 Array"
      CaptionAlign    =   1
      FocusStyle      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmThemedButton.frx":28424
      OptionButtonMultiSelect=   -1  'True
      OverColor       =   -2147483633
      UseParentBackColor=   -1  'True
      Value           =   0
   End
   Begin VB.Image imgDemo 
      Height          =   420
      Left            =   2640
      Picture         =   "frmThemedButton.frx":289BE
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2052
   End
End
Attribute VB_Name = "frmThemedButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub thbQuit_Click()

   Unload Me

End Sub
