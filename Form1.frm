VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings..."
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.HeaderPicture HeaderPicture8 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      Caption         =   "Green?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      Gradient        =   1
      GradientStart   =   8454016
      GradientFinish  =   16384
   End
   Begin Project1.HeaderPicture HeaderPicture7 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      Caption         =   "Rounded Top"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      GradientStart   =   8421504
      Shape           =   2
   End
   Begin Project1.HeaderPicture HeaderPicture6 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      Caption         =   "Vertical Gradient"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      FontColor       =   12582912
      Gradient        =   1
      GradientStart   =   16777215
      GradientFinish  =   14378786
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DB6722&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Text            =   "Use me as a Frame also!"
      Top             =   3240
      Width           =   3855
   End
   Begin Project1.HeaderPicture HeaderPicture3 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      Caption         =   "Rounded Style"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      GradientStart   =   16094834
      GradientFinish  =   16711680
   End
   Begin Project1.HeaderPicture HeaderPicture1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      Caption         =   "Opaque Look"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      GradientStart   =   16094834
      GradientFinishStyle=   1
   End
   Begin Project1.HeaderPicture HeaderPicture4 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      Gradient        =   1
      GradientStart   =   15188135
      GradientFinish  =   14378786
      Shape           =   0
   End
   Begin Project1.HeaderPicture HeaderPicture5 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      FontColor       =   14378786
      Gradient        =   1
      GradientStart   =   16777215
      GradientFinish  =   14378786
      Shape           =   2
   End
   Begin Project1.HeaderPicture HeaderPicture2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   8
      FontColor       =   4210752
      Gradient        =   1
      GradientStart   =   16777215
      GradientFinish  =   8421504
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
