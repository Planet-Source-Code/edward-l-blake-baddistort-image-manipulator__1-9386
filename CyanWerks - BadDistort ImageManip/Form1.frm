VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edward L. Blake's BadDistort Image Manipulator"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   DrawMode        =   14  'Copy Pen
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "D1\D2 Pair"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   60
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Miscelaneous"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   7680
      TabIndex        =   41
      Top             =   120
      Width           =   2535
      Begin VB.Frame Frame7 
         Caption         =   """Cartoonizer""\Posterize"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   62
         Top             =   4200
         Width           =   2295
         Begin VB.HScrollBar HScroll22 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   255
            TabIndex        =   65
            Top             =   600
            Width           =   1335
         End
         Begin VB.HScrollBar HScroll21 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   255
            Min             =   127
            TabIndex        =   63
            Top             =   240
            Value           =   255
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Amp:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   66
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Cutoff:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Distortion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   2295
         Begin VB.CheckBox Check2 
            Caption         =   "Monochrome"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   61
            Top             =   1680
            Width           =   1305
         End
         Begin VB.HScrollBar HScroll20 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   255
            TabIndex        =   56
            Top             =   1320
            Width           =   1335
         End
         Begin VB.HScrollBar HScroll19 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   255
            TabIndex        =   55
            Top             =   960
            Width           =   1335
         End
         Begin VB.HScrollBar HScroll18 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   255
            TabIndex        =   54
            Top             =   600
            Width           =   1335
         End
         Begin VB.HScrollBar HScroll17 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   255
            TabIndex        =   53
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Master"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   59
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "B-Amp:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "G-Amp:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "R-Amp:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Brightness"
         Height          =   1695
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   2295
         Begin VB.HScrollBar HScroll16 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   512
            TabIndex        =   50
            Top             =   1320
            Value           =   255
            Width           =   1335
         End
         Begin VB.HScrollBar HScroll15 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   512
            TabIndex        =   48
            Top             =   960
            Value           =   255
            Width           =   1335
         End
         Begin VB.HScrollBar HScroll14 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   512
            TabIndex        =   47
            Top             =   600
            Value           =   255
            Width           =   1335
         End
         Begin VB.HScrollBar HScroll13 
            Height          =   255
            LargeChange     =   64
            Left            =   840
            Max             =   512
            TabIndex        =   46
            Top             =   240
            Value           =   255
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Master"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "B-Bright"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "G-Bright"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "R-Bright"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tuning Variables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   3255
      Begin VB.Frame Frame3 
         Caption         =   "D2 Pixel Oscillator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   3015
         Begin VB.HScrollBar HScroll12 
            Height          =   255
            LargeChange     =   5
            Left            =   840
            Max             =   10
            Min             =   1
            TabIndex        =   28
            Top             =   2040
            Value           =   1
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll11 
            Height          =   255
            LargeChange     =   5
            Left            =   840
            Max             =   10
            Min             =   1
            TabIndex        =   27
            Top             =   1680
            Value           =   1
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll10 
            Height          =   255
            LargeChange     =   5
            Left            =   840
            Max             =   10
            Min             =   1
            TabIndex        =   26
            Top             =   1320
            Value           =   1
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll6 
            Height          =   255
            LargeChange     =   50
            Left            =   840
            Max             =   100
            Min             =   -100
            TabIndex        =   16
            Top             =   960
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll5 
            Height          =   255
            LargeChange     =   50
            Left            =   840
            Max             =   100
            Min             =   -100
            TabIndex        =   15
            Top             =   600
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll4 
            Height          =   255
            LargeChange     =   50
            Left            =   840
            Max             =   100
            Min             =   -100
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   11
            Left            =   2520
            TabIndex        =   40
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   39
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   9
            Left            =   2520
            TabIndex        =   38
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   34
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   33
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   32
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "B-OSC-F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Blue Oscillator Frequency"
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "G-OSC-F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Green Oscillator Frequency"
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "R-OSC-F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Red Oscillator Frequency"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "B-OSC-A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Blue Oscillator Amplitude"
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "G-OSC-A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Green Oscillator Amplitude"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "R-OSC-A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Red Oscillator Amplitude"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "D1 Pixel Oscillator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3015
         Begin VB.HScrollBar HScroll9 
            Height          =   255
            LargeChange     =   5
            Left            =   840
            Max             =   10
            Min             =   1
            TabIndex        =   22
            Top             =   2040
            Value           =   1
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll8 
            Height          =   255
            LargeChange     =   5
            Left            =   840
            Max             =   10
            Min             =   1
            TabIndex        =   21
            Top             =   1680
            Value           =   1
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll7 
            Height          =   255
            LargeChange     =   5
            Left            =   840
            Max             =   10
            Min             =   1
            TabIndex        =   20
            Top             =   1320
            Value           =   1
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            LargeChange     =   50
            Left            =   840
            Max             =   100
            Min             =   -100
            TabIndex        =   10
            Top             =   960
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            LargeChange     =   50
            Left            =   840
            Max             =   100
            Min             =   -100
            TabIndex        =   9
            Top             =   600
            Width           =   1575
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            LargeChange     =   50
            Left            =   840
            Max             =   100
            Min             =   -100
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   8
            Left            =   2520
            TabIndex        =   37
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   7
            Left            =   2520
            TabIndex        =   36
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   6
            Left            =   2520
            TabIndex        =   35
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   31
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   30
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   29
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "B-OSC-F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Blue Oscillator Frequency"
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "G-OSC-F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Green Oscillator Frequency"
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "R-OSC-F"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Red Oscillator Frequency"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "B-OSC-A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Blue Oscillator Amplitude"
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "G-OSC-A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Green Oscilator Amplitude"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "R-OSC-A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Red Oscilator Amplitude"
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   1
      Top             =   2760
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   277
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 2000 Edward L. Blake
' Email: blakee@cyanwerks.com
'        blakee@rovoscape.com
'
' This somewhat strange project demonstrates using the
' GetPixel and SetPixel functions to separate the channels
' and process an image as it's being drawn. Different values
' can be tweeked as the image is continually filtered. This
' particular project can distort images with simple filters
' (Pixel Intensity Oscillation, Channel separated Brightness,
' both monochrome and non-monochrome Noise Distortion,
' Cartoonization\Posterize).
'
' http://www.cyanwerks.com/
'

'Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Dim GetOut As Boolean
Dim NotYet As Boolean

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        Frame3.Visible = False
    Else
        Frame3.Visible = True
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        HScroll17.Visible = False
        HScroll18.Visible = False
        HScroll19.Visible = False
    Else
        HScroll17.Visible = True
        HScroll18.Visible = True
        HScroll19.Visible = True
    End If
End Sub

Private Sub Form_Activate()
    If Not NotYet Then DoFrames
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    GetOut = True
End Sub

Private Sub DoFrames()
    Dim Ix As Long, Iy As Long
    Dim Gr As Long
    Dim P1H As Long, P2H As Long
    Dim D1 As Long, D2 As Long
    Dim MonoDst As Boolean
    Dim DPair As Boolean
    Dim D1ROA As Integer
    Dim D1GOA As Integer
    Dim D1BOA As Integer
    Dim D1ROF As Integer
    Dim D1GOF As Integer
    Dim D1BOF As Integer
    Dim D2ROA As Integer
    Dim D2GOA As Integer
    Dim D2BOA As Integer
    Dim D2ROF As Integer
    Dim D2GOF As Integer
    Dim D2BOF As Integer
    Dim RBr As Integer
    Dim GBr As Integer
    Dim BBr As Integer
    Dim RDst As Integer
    Dim GDst As Integer
    Dim BDst As Integer
    Dim RPx As Integer
    Dim GPx As Integer
    Dim BPx As Integer
    Dim Cct As Integer
    Dim CAm As Integer
    
    NotYet = True
    With Picture1
        Do While True
            P1H = Picture1.hdc
            P2H = Picture2.hdc
            
            For Iy = 0 To 200
                DoEvents
                
                If GetOut = False Then
                    D1ROA = HScroll1.Value
                    D1GOA = HScroll2.Value
                    D1BOA = HScroll3.Value
                    D1ROF = HScroll7.Value
                    D1GOF = HScroll8.Value
                    D1BOF = HScroll9.Value
                    D2ROA = HScroll4.Value
                    D2GOA = HScroll5.Value
                    D2BOA = HScroll6.Value
                    D2ROF = HScroll10.Value
                    D2GOF = HScroll11.Value
                    D2BOF = HScroll12.Value
                    RBr = HScroll13.Value
                    GBr = HScroll14.Value
                    BBr = HScroll15.Value
                    
                    Cct = HScroll21.Value
                    CAm = HScroll22.Value
                    
                    If Check2.Value = 0 Then
                        MonoDst = False
                        RDst = HScroll17.Value
                        GDst = HScroll18.Value
                        BDst = HScroll19.Value
                    Else
                        MonoDst = True
                        RDst = HScroll20.Value
                    End If
                    If Check1.Value = 0 Then DPair = False Else DPair = True
                End If
                
                If DPair Then
                    For Ix = 0 To 320 Step 2
                        If GetOut Then Exit Do
                        Gr = GetPixel(P1H, Ix, Iy)
                        
                        If MonoDst Then
                            RPx = Red(Gr)
                            GPx = Green(Gr)
                            BPx = Blue(Gr)
                            If RPx > Cct Then RPx = CAm
                            If GPx > Cct Then GPx = CAm
                            If BPx > Cct Then BPx = CAm
                            If RPx < (255 - Cct) Then RPx = 255 - CAm
                            If GPx < (255 - Cct) Then GPx = 255 - CAm
                            If BPx < (255 - Cct) Then BPx = 255 - CAm
                            GDst = (Rnd * RDst - RDst / 2)
                            D1 = RGB(Om((RPx + D1ROA * Sin(Ix + Iy / D1ROF)) / 255 * RBr + GDst), Om((GPx + D1GOA * Sin(Ix + Iy / D1GOF)) / 255 * GBr + GDst), Om((BPx + D1BOA * Sin(Ix + Iy / D1BOF)) / 255 * BBr + GDst))
                            D2 = RGB(Om((RPx - 40 + D2ROA * Sin(Ix + Iy / D2ROF)) / 255 * RBr + GDst), Om((GPx - 40 + D2GOA * Sin(Ix + Iy / D2GOF)) / 255 * GBr + GDst), Om((BPx - 40 + D2BOA * Sin(Ix + Iy / D2BOF)) / 255 * BBr + GDst))
                        Else
                            RPx = Red(Gr)
                            GPx = Green(Gr)
                            BPx = Blue(Gr)
                            If RPx > Cct Then RPx = CAm
                            If GPx > Cct Then GPx = CAm
                            If BPx > Cct Then BPx = CAm
                            If RPx < (255 - Cct) Then RPx = 255 - CAm
                            If GPx < (255 - Cct) Then GPx = 255 - CAm
                            If BPx < (255 - Cct) Then BPx = 255 - CAm
                            D1 = RGB(Om((RPx + D1ROA * Sin(Ix + Iy / D1ROF)) / 255 * RBr + (Rnd * RDst - RDst / 2)), Om((GPx + D1GOA * Sin(Ix + Iy / D1GOF)) / 255 * GBr + (Rnd * GDst - GDst / 2)), Om((BPx + D1BOA * Sin(Ix + Iy / D1BOF)) / 255 * BBr + (Rnd * BDst - BDst / 2)))
                            D2 = RGB(Om((RPx - 40 + D2ROA * Sin(Ix + Iy / D2ROF)) / 255 * RBr + (Rnd * RDst - RDst / 2)), Om((GPx - 40 + D2GOA * Sin(Ix + Iy / D2GOF)) / 255 * GBr + (Rnd * GDst - GDst / 2)), Om((BPx - 40 + D2BOA * Sin(Ix + Iy / D2BOF)) / 255 * BBr + (Rnd * BDst - BDst / 2)))
                        End If
                        SetPixel P2H, Ix - 1, Iy - 1, D1
                        SetPixel P2H, Ix, Iy - 1, D2
                        SetPixel P2H, Ix + 1, Iy - 1, D1
                        
                        SetPixel P2H, Ix - 1, Iy, D2
                        SetPixel P2H, Ix, Iy, Gr
                        SetPixel P2H, Ix + 1, Iy, D2
                        
                        SetPixel P2H, Ix - 1, Iy + 1, D1
                        SetPixel P2H, Ix, Iy + 1, D2
                        SetPixel P2H, Ix + 1, Iy + 1, D1
                    Next Ix
                Else
                    For Ix = 0 To 320
                        If GetOut Then Exit Do
                        Gr = GetPixel(P1H, Ix, Iy)
                        
                        If MonoDst Then
                            RPx = Red(Gr)
                            GPx = Green(Gr)
                            BPx = Blue(Gr)
                            If RPx > Cct Then RPx = CAm
                            If GPx > Cct Then GPx = CAm
                            If BPx > Cct Then BPx = CAm
                            If RPx < (255 - Cct) Then RPx = 255 - CAm
                            If GPx < (255 - Cct) Then GPx = 255 - CAm
                            If BPx < (255 - Cct) Then BPx = 255 - CAm
                            GDst = (Rnd * RDst - RDst / 2)
                            D1 = RGB(Om((RPx + D1ROA * Sin(Ix + Iy / D1ROF)) / 255 * RBr + GDst), Om((GPx + D1GOA * Sin(Ix + Iy / D1GOF)) / 255 * GBr + GDst), Om((BPx + D1BOA * Sin(Ix + Iy / D1BOF)) / 255 * BBr + GDst))
                            D2 = RGB(Om((RPx - 40 + D2ROA * Sin(Ix + Iy / D2ROF)) / 255 * RBr + GDst), Om((GPx - 40 + D2GOA * Sin(Ix + Iy / D2GOF)) / 255 * GBr + GDst), Om((BPx - 40 + D2BOA * Sin(Ix + Iy / D2BOF)) / 255 * BBr + GDst))
                        Else
                            RPx = Red(Gr)
                            GPx = Green(Gr)
                            BPx = Blue(Gr)
                            If RPx > Cct Then RPx = CAm
                            If GPx > Cct Then GPx = CAm
                            If BPx > Cct Then BPx = CAm
                            If RPx < (255 - Cct) Then RPx = 255 - CAm
                            If GPx < (255 - Cct) Then GPx = 255 - CAm
                            If BPx < (255 - Cct) Then BPx = 255 - CAm
                            D1 = RGB(Om((RPx + D1ROA * Sin(Ix + Iy / D1ROF)) / 255 * RBr + (Rnd * RDst - RDst / 2)), Om((GPx + D1GOA * Sin(Ix + Iy / D1GOF)) / 255 * GBr + (Rnd * GDst - GDst / 2)), Om((BPx + D1BOA * Sin(Ix + Iy / D1BOF)) / 255 * BBr + (Rnd * BDst - BDst / 2)))
                            D2 = RGB(Om((RPx - 40 + D2ROA * Sin(Ix + Iy / D2ROF)) / 255 * RBr + (Rnd * RDst - RDst / 2)), Om((GPx - 40 + D2GOA * Sin(Ix + Iy / D2GOF)) / 255 * GBr + (Rnd * GDst - GDst / 2)), Om((BPx - 40 + D2BOA * Sin(Ix + Iy / D2BOF)) / 255 * BBr + (Rnd * BDst - BDst / 2)))
                        End If
                        
                        SetPixel P2H, Ix - 1, Iy - 1, D1
                        SetPixel P2H, Ix, Iy - 1, D2
                        SetPixel P2H, Ix + 1, Iy - 1, D1
                        
                        SetPixel P2H, Ix - 1, Iy, D2
                        SetPixel P2H, Ix, Iy, Gr
                        SetPixel P2H, Ix + 1, Iy, D2
                        
                        SetPixel P2H, Ix - 1, Iy + 1, D1
                        SetPixel P2H, Ix, Iy + 1, D2
                        SetPixel P2H, Ix + 1, Iy + 1, D1
                    Next Ix
                End If
                
            Next Iy
'            Picture2.AutoRedraw = True
'            Picture1.PaintPicture Picture2.Image, 0, 0, 100, 100
'            Picture2.AutoRedraw = False
        Loop
    End With
End Sub

Private Sub HScroll1_Change()
    Label13(0).Caption = HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
    Label13(0).Caption = HScroll1.Value
End Sub

Private Sub HScroll10_Change()
    Label13(9).Caption = HScroll10.Value
End Sub
Private Sub HScroll10_Scroll()
    Label13(9).Caption = HScroll10.Value
End Sub

Private Sub HScroll11_Change()
    Label13(10).Caption = HScroll11.Value
End Sub
Private Sub HScroll11_Scroll()
    Label13(10).Caption = HScroll11.Value
End Sub

Private Sub HScroll12_Change()
    Label13(11).Caption = HScroll12.Value
End Sub
Private Sub HScroll12_Scroll()
    Label13(11).Caption = HScroll12.Value
End Sub

Private Sub HScroll16_Change()
    HScroll13.Value = HScroll16.Value
    HScroll14.Value = HScroll16.Value
    HScroll15.Value = HScroll16.Value
End Sub

Private Sub HScroll16_Scroll()
    HScroll13.Value = HScroll16.Value
    HScroll14.Value = HScroll16.Value
    HScroll15.Value = HScroll16.Value
End Sub

Private Sub HScroll2_Change()
    Label13(1).Caption = HScroll2.Value
End Sub
Private Sub HScroll2_Scroll()
    Label13(1).Caption = HScroll2.Value
End Sub

Private Sub HScroll20_Change()
    HScroll17.Value = HScroll20.Value
    HScroll18.Value = HScroll20.Value
    HScroll19.Value = HScroll20.Value
End Sub

Private Sub HScroll20_Scroll()
    HScroll17.Value = HScroll20.Value
    HScroll18.Value = HScroll20.Value
    HScroll19.Value = HScroll20.Value
End Sub

Private Sub HScroll3_Change()
    Label13(2).Caption = HScroll3.Value
End Sub
Private Sub HScroll3_Scroll()
    Label13(2).Caption = HScroll3.Value
End Sub

Private Sub HScroll4_Change()
    Label13(3).Caption = HScroll4.Value
End Sub
Private Sub HScroll4_Scroll()
    Label13(3).Caption = HScroll4.Value
End Sub

Private Sub HScroll5_Change()
    Label13(4).Caption = HScroll5.Value
End Sub
Private Sub HScroll5_Scroll()
    Label13(4).Caption = HScroll5.Value
End Sub

Private Sub HScroll6_Change()
    Label13(5).Caption = HScroll6.Value
End Sub
Private Sub HScroll6_Scroll()
    Label13(5).Caption = HScroll6.Value
End Sub

Private Sub HScroll7_Change()
    Label13(6).Caption = HScroll7.Value
End Sub
Private Sub HScroll7_Scroll()
    Label13(6).Caption = HScroll7.Value
End Sub

Private Sub HScroll8_Change()
    Label13(7).Caption = HScroll8.Value
End Sub
Private Sub HScroll8_Scroll()
    Label13(7).Caption = HScroll8.Value
End Sub

Private Sub HScroll9_Change()
    Label13(8).Caption = HScroll9.Value
End Sub
Private Sub HScroll9_Scroll()
    Label13(8).Caption = HScroll9.Value
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'SetPixel Picture1.hdc, x, y, vbWhite
End Sub

Function Red(ByVal Color As Long)
Red = Int(Color - Int((Color - Int(Color / 65536) * 65536) / 256) * 256 - Int(Color / 65536) * 65536)
End Function
Function Green(ByVal Color As Long)
Green = Int((Color - Int(Color / 65536) * 65536) / 256)
End Function
Function Blue(ByVal Color As Long)
Blue = Int(Color / 65536)
End Function

Function Om(ByVal n As Long)
    If n < 0 Then Om = 0 Else Om = n
End Function

