VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HOVERBUTTON.OCX"
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   2820
   ClientLeft      =   8100
   ClientTop       =   4620
   ClientWidth     =   3270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   3270
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "0"
      CaptionDown     =   "0"
      CaptionOver     =   "0"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.TextBox CalcDisPlay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "1"
      CaptionDown     =   "1"
      CaptionOver     =   "1"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "3"
      CaptionDown     =   "3"
      CaptionOver     =   "3"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "2"
      CaptionDown     =   "2"
      CaptionOver     =   "2"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "4"
      CaptionDown     =   "4"
      CaptionOver     =   "4"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "5"
      CaptionDown     =   "5"
      CaptionOver     =   "5"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   7
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "6"
      CaptionDown     =   "6"
      CaptionOver     =   "6"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "7"
      CaptionDown     =   "7"
      CaptionOver     =   "7"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "8"
      CaptionDown     =   "8"
      CaptionOver     =   "8"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   9
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "9"
      CaptionDown     =   "9"
      CaptionOver     =   "9"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   10
      Left            =   840
      TabIndex        =   11
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "."
      CaptionDown     =   "."
      CaptionOver     =   "."
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   11
      Left            =   1680
      TabIndex        =   12
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "="
      CaptionDown     =   "="
      CaptionOver     =   "="
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   12
      Left            =   2520
      TabIndex        =   13
      Top             =   960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "/"
      CaptionDown     =   "/"
      CaptionOver     =   "/"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   13
      Left            =   2520
      TabIndex        =   14
      Top             =   1440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "*"
      CaptionDown     =   "*"
      CaptionOver     =   "*"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   14
      Left            =   2520
      TabIndex        =   15
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "+"
      CaptionDown     =   "+"
      CaptionOver     =   "+"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   15
      Left            =   2520
      TabIndex        =   16
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "-"
      CaptionDown     =   "-"
      CaptionOver     =   "-"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483633
      Border          =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   14737632
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483628
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "C"
      CaptionDown     =   "C"
      CaptionOver     =   "C"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FirstValue As String
Dim SecValue As String
Dim Result As Double
Dim Sign As String
Dim PointPress As Integer
Dim resetdisplay As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    Const conZero As Integer = 48, conNine As Integer = 57, conEqual As Integer = 61
    Const conDivide As Integer = 47, conMultiply As Integer = 42, conAdd As Integer = 43
    Const conSubtract As Integer = 45, conReturn As Integer = 13, conPoint As Integer = 46
    If (KeyAscii < conZero Or KeyAscii > conNine) And KeyAscii <> conEqual _
        And KeyAscii <> conDivide And KeyAscii <> conMultiply And KeyAscii <> conSubtract _
        And KeyAscii <> conAdd And KeyAscii <> conReturn And KeyAscii <> conPoint And KeyAscii = "67" And KeyAscii = "99" Then
            KeyAscii = 0
    Else
        
        Select Case KeyAscii
        Case "99": CalcReset
        Case "67": CalcReset
        Case "49": ShowDisplay ("1")
        Case "50": ShowDisplay ("2")
        Case "51": ShowDisplay ("3")
        Case "52": ShowDisplay ("4")
        Case "53": ShowDisplay ("5")
        Case "54": ShowDisplay ("6")
        Case "55": ShowDisplay ("7")
        Case "56": ShowDisplay ("8")
        Case "57": ShowDisplay ("9")
        Case "48":
                    If Not (CalcDisPlay.Text = "0.") Or (PointPress = 1) Then
                        ShowDisplay ("0")
                    End If
        Case conDivide
            KeyAscii = 0
            CheckSign ("/")
        
        Case conMultiply
            KeyAscii = 0
            CheckSign ("*")
        
        Case conSubtract
            KeyAscii = 0
            CheckSign ("-")
        Case conAdd
            KeyAscii = 0
            CheckSign ("+")
        Case conEqual
            KeyAscii = 0
            CheckSign ("=")
            
        Case conReturn
            KeyAscii = 0
            CheckSign ("=")
            
        Case conPoint
            KeyAscii = 0
            If (InStr(2, StrReverse(CalcDisPlay.Text), ".") < 1) And (CalcDisPlay <> "0.") Then
                ShowDisplay (".")
             Else
                PointPress = 1
             End If
    
    End Select
    End If
     
End Sub


Private Sub Command1_Click(Index As Integer)
Dim ButtonSelect As String
ButtonSelect = Command1.Item(Index).Caption
Select Case ButtonSelect
    Case "0":
             If Not (CalcDisPlay.Text = "0.") Or (PointPress = 1) Then
                ShowDisplay (ButtonSelect)
             End If
    Case "1": ShowDisplay (ButtonSelect)
    Case "2": ShowDisplay (ButtonSelect)
    Case "3": ShowDisplay (ButtonSelect)
    Case "4": ShowDisplay (ButtonSelect)
    Case "5": ShowDisplay (ButtonSelect)
    Case "6": ShowDisplay (ButtonSelect)
    Case "7": ShowDisplay (ButtonSelect)
    Case "8": ShowDisplay (ButtonSelect)
    Case "9": ShowDisplay (ButtonSelect)
    Case "0": ShowDisplay (ButtonSelect)
    Case ".":

             If (InStr(2, StrReverse(CalcDisPlay.Text), ".") < 1) And (CalcDisPlay <> "0.") Then
                ShowDisplay (ButtonSelect)
             Else
                PointPress = 1
             End If
    Case "C": CalcReset
    Case "/": CheckSign (ButtonSelect)
    Case "*": CheckSign (ButtonSelect)
    Case "+": CheckSign (ButtonSelect)
    Case "-": CheckSign (ButtonSelect)
    Case "=": CheckSign (ButtonSelect)
End Select

End Sub
Function CheckSign(ButtonPress As String)
If ButtonPress = "=" Then
    ButtonPress = ""
End If
    If FirstValue = "" Then
        FirstValue = CalcDisPlay.Text
        resetdisplay = 1
        PointPress = 0
        Sign = ButtonPress
    ElseIf SecValue = "" Then
            If FirstValue <> "" And SecValue = "" And Sign = "" Then
                resetdisplay = 1
                Sign = ButtonPress
            Else
                SecValue = CalcDisPlay.Text
                resetdisplay = 1
            End If
    End If
    If SecValue <> "" Then
            ComputeNumber
            Sign = ButtonPress
    End If
End Function
Function ComputeNumber()
    If (FirstValue = "" Or FirstValue = "0.") Then
       FirstValue = "0"
    End If
    If (SecValue = "" Or SecValue = "0.") Then
       SecValue = "0"
    End If
    
    Select Case Sign
        Case "+": Result = CDbl(FirstValue) + CDbl(SecValue)
        Case "-": Result = CDbl(FirstValue) - CDbl(SecValue)
        Case "*": Result = CDbl(FirstValue) * CDbl(SecValue)
        Case "/":
                If (SecValue = "" Or SecValue = "0") Then
                   MsgBox "Cant divide by zero."
                Else
                   Result = CDbl(FirstValue) / CDbl(SecValue)
                End If
    End Select
    PointPress = 0
    CalcDisPlay.Text = Result
    FirstValue = CalcDisPlay.Text
    Sign = ""
    SecValue = ""
    'Command1(11).SetFocus
    
End Function
Function ShowDisplay(ButtonPress As String)
        If Sign = "" Then
            FirstValue = ""
            SecValue = ""
        End If
        If resetdisplay = 1 Then
            resetdisplay = 0
            CalcDisPlay.Text = ""
        End If
           If PointPress = 0 Then
                PointPress = 1
                CalcDisPlay.Text = ButtonPress
           Else
                CalcDisPlay.Text = CalcDisPlay.Text & ButtonPress
           End If
        'Command1(11).SetFocus
End Function
Private Sub CalcReset()
    CalcDisPlay.Text = "0."
    PointPress = 0
    FirstValue = ""
    SecValue = ""
    Sign = ""
    resetdisplay = 0
   
    
End Sub
Sub Form_Load()

    CalcReset
    KeyPreview = True
End Sub



