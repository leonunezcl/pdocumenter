VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Documenter"
   ClientHeight    =   5895
   ClientLeft      =   1830
   ClientTop       =   2040
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   131
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAbrir"
            Object.ToolTipText     =   "Abrir archivo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdSelTodo"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cmdLimTodo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPreview"
            Object.ToolTipText     =   "Visualizar informe"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdImprimir"
            Object.ToolTipText     =   "Imprimir archivo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdWeb"
            Object.ToolTipText     =   "Ir al sitio web de vbsoftware"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSalir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   4035
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":030A
            Key             =   ""
            Object.Tag             =   "&Abrir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":041E
            Key             =   ""
            Object.Tag             =   "&Imprimir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0532
            Key             =   ""
            Object.Tag             =   "&Visualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0D92
            Key             =   ""
            Object.Tag             =   "&Indice"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":10AE
            Key             =   ""
            Object.Tag             =   "&Salir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":150A
            Key             =   ""
            Object.Tag             =   "&Limpiar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1DE6
            Key             =   ""
            Object.Tag             =   "&Seleccionar"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5190
      Left            =   15
      ScaleHeight     =   344
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   130
      Top             =   375
      Width           =   360
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3720
      ScaleHeight     =   285
      ScaleWidth      =   480
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4260
      ScaleHeight     =   5.027
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   8.467
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   5310
      Sorted          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   4815
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   5310
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   780
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5190
      Left            =   390
      TabIndex        =   11
      Top             =   375
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   9155
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "&Documentacion"
      TabPicture(0)   =   "Main.frx":26C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdPrintSetup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdView"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPrint"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSelectAll"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdClear"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdHelp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Pre&ferencias"
      TabPicture(1)   =   "Main.frx":26DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame(3)"
      Tab(1).Control(1)=   "TabOptions"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Archivos y Procedimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4065
         Left            =   105
         TabIndex        =   92
         Top             =   1020
         Width           =   3330
         Begin MSOutl.Outline Outline 
            Height          =   3765
            Left            =   75
            TabIndex        =   3
            Top             =   180
            Width           =   3180
            _Version        =   65536
            _ExtentX        =   5609
            _ExtentY        =   6641
            _StockProps     =   77
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            PathSeparator   =   "-"
            PicturePlus     =   "Main.frx":26FA
            PictureMinus    =   "Main.frx":27F4
            PictureLeaf     =   "Main.frx":28EE
            PictureOpen     =   "Main.frx":31C8
            PictureClosed   =   "Main.frx":3AA2
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "&Seleccionar archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   7
         Left            =   105
         TabIndex        =   78
         Top             =   345
         Width           =   7425
         Begin VB.CommandButton cmdPrevFile 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7005
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdPickFile 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6750
            TabIndex        =   1
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtProject 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   210
            Width           =   7170
         End
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Ayuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4755
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   4320
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Frame Frame 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Index           =   2
         Left            =   3525
         TabIndex        =   75
         Top             =   2490
         Width           =   4005
         Begin VB.CheckBox chkPreview 
            Caption         =   "Previe&w"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1890
            TabIndex        =   8
            Top             =   990
            Width           =   930
         End
         Begin VB.Label lblViewSize 
            Alignment       =   1  'Right Justify
            Caption         =   "(Zoom 100%)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2850
            TabIndex        =   89
            Top             =   990
            Width           =   1035
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   9
            X1              =   1755
            X2              =   1755
            Y1              =   1260
            Y2              =   900
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   8
            X1              =   1770
            X2              =   1770
            Y1              =   900
            Y2              =   1275
         End
         Begin VB.Label lblPrinted 
            AutoSize        =   -1  'True
            Caption         =   "Ninguna"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1095
            TabIndex        =   80
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Paginas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   79
            Top             =   990
            UseMnemonic     =   0   'False
            Width           =   615
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   15
            X2              =   4000
            Y1              =   885
            Y2              =   885
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   30
            X2              =   3980
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Label lblPrinter 
            Caption         =   "Ninguna seleccionada -  Seleccione 'Configurar Impresora'"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   76
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   3780
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Archivos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Index           =   0
         Left            =   3525
         TabIndex        =   70
         Top             =   1020
         Width           =   4005
         Begin VB.Label lblSelProcs 
            AutoSize        =   -1  'True
            Caption         =   "Ninguno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3090
            TabIndex        =   96
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   720
         End
         Begin VB.Label lblProcedures 
            AutoSize        =   -1  'True
            Caption         =   "Ninguno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1275
            TabIndex        =   95
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Seleccionado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   1995
            TabIndex        =   94
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   975
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Procedimientos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   93
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblName 
            Caption         =   "(No hay item seleccionado)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   82
            Top             =   870
            Width           =   3780
         End
         Begin VB.Label lblType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   81
            Top             =   1125
            UseMnemonic     =   0   'False
            Width           =   3780
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   30
            X2              =   3980
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   15
            X2              =   4000
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Label lblFiles 
            AutoSize        =   -1  'True
            Caption         =   "Ninguno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1275
            TabIndex        =   74
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Seleccionado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1995
            TabIndex        =   73
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label 
            Caption         =   "Archivos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   915
         End
         Begin VB.Label lblSelFiles 
            AutoSize        =   -1  'True
            Caption         =   "Ninguno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3090
            TabIndex        =   71
            Top             =   225
            UseMnemonic     =   0   'False
            Width           =   720
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Pagina de Ejemplo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Index           =   3
         Left            =   -74895
         TabIndex        =   63
         Top             =   360
         Width           =   3735
         Begin VB.Timer TmrPaint 
            Enabled         =   0   'False
            Interval        =   2000
            Left            =   3195
            Top             =   3135
         End
         Begin VB.PictureBox picPage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2865
            Left            =   165
            ScaleHeight     =   2835
            ScaleWidth      =   2070
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   255
            Width           =   2100
         End
         Begin ComctlLib.Slider sldRight 
            Height          =   225
            Left            =   1440
            TabIndex        =   27
            Top             =   3120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   397
            _Version        =   327682
            Max             =   70
            SelStart        =   70
            TickFrequency   =   10
            Value           =   70
         End
         Begin ComctlLib.Slider sldBottom 
            Height          =   1185
            Left            =   2265
            TabIndex        =   26
            Top             =   2040
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   2090
            _Version        =   327682
            Orientation     =   1
            Max             =   99
            SelStart        =   99
            TickFrequency   =   10
            Value           =   99
         End
         Begin ComctlLib.Slider sldTop 
            Height          =   1185
            Left            =   2265
            TabIndex        =   25
            Top             =   165
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   2090
            _Version        =   327682
            Orientation     =   1
            Max             =   99
            TickFrequency   =   10
         End
         Begin ComctlLib.Slider sldLeft 
            Height          =   225
            Left            =   60
            TabIndex        =   28
            Top             =   3120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   397
            _Version        =   327682
            Max             =   70
            TickFrequency   =   10
         End
         Begin VB.Label Label 
            Caption         =   "Orientacion:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   90
            TabIndex        =   86
            Top             =   3720
            UseMnemonic     =   0   'False
            Width           =   900
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tamaño pagina:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   85
            Top             =   3960
            UseMnemonic     =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblOrient 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   990
            TabIndex        =   84
            Top             =   3720
            UseMnemonic     =   0   'False
            Width           =   2610
         End
         Begin VB.Label lblSize 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1320
            TabIndex        =   83
            Top             =   3960
            UseMnemonic     =   0   'False
            Width           =   2355
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   30
            X2              =   3720
            Y1              =   3615
            Y2              =   3615
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            X1              =   15
            X2              =   3720
            Y1              =   3630
            Y2              =   3630
         End
         Begin VB.Label lblMM 
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2565
            TabIndex        =   69
            Top             =   3390
            Width           =   255
         End
         Begin VB.Label lblLeft 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   68
            Top             =   3390
            Width           =   330
         End
         Begin VB.Label lblRight 
            Alignment       =   1  'Right Justify
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   1965
            TabIndex        =   67
            Top             =   3390
            Width           =   330
         End
         Begin VB.Label lblBottom 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2565
            TabIndex        =   66
            Top             =   2970
            Width           =   330
         End
         Begin VB.Label lblTop 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2565
            TabIndex        =   65
            Top             =   240
            Width           =   330
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Limpiar Todo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         TabIndex        =   5
         Top             =   4320
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Seleccionar Todo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   4
         Top             =   4320
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6450
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4320
         Visible         =   0   'False
         Width           =   1110
      End
      Begin TabDlg.SSTab TabOptions 
         Height          =   4155
         Left            =   -71040
         TabIndex        =   12
         Top             =   465
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   7329
         _Version        =   393216
         TabOrientation  =   3
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   6
         TabHeight       =   520
         WordWrap        =   0   'False
         TabCaption(0)   =   "Opciones"
         TabPicture(0)   =   "Main.frx":437C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "chkSortIndex"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkIndex"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkControlPage"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkProcNames"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "chkProject"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkSeparator"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "chkCode"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "chkControlNames"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "chkSortControls"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "chkIcon"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkProcPage"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "chkFormIcons"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Pagina"
         TabPicture(1)   =   "Main.frx":4398
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label(8)"
         Tab(1).Control(1)=   "Line(16)"
         Tab(1).Control(2)=   "Line(17)"
         Tab(1).Control(3)=   "txtOwner(0)"
         Tab(1).Control(4)=   "txtOwner(1)"
         Tab(1).Control(5)=   "chkPageNumbers"
         Tab(1).Control(6)=   "chkHeader"
         Tab(1).Control(7)=   "chkDate"
         Tab(1).Control(8)=   "chkTime"
         Tab(1).Control(9)=   "chkResetPage"
         Tab(1).Control(10)=   "chkFooter"
         Tab(1).Control(11)=   "optPagePos(0)"
         Tab(1).Control(12)=   "optPagePos(1)"
         Tab(1).ControlCount=   13
         TabCaption(2)   =   "Usuario"
         TabPicture(2)   =   "Main.frx":43B4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Line(10)"
         Tab(2).Control(1)=   "Line(11)"
         Tab(2).Control(2)=   "lblExtention"
         Tab(2).Control(3)=   "Line(14)"
         Tab(2).Control(4)=   "Line(15)"
         Tab(2).Control(5)=   "lblZoomDialog"
         Tab(2).Control(6)=   "lblSoundFile"
         Tab(2).Control(7)=   "Line(18)"
         Tab(2).Control(8)=   "Line(19)"
         Tab(2).Control(9)=   "Line(20)"
         Tab(2).Control(10)=   "Line(21)"
         Tab(2).Control(11)=   "chkPlayWaves"
         Tab(2).Control(12)=   "cboExtention"
         Tab(2).Control(13)=   "cboZoom"
         Tab(2).Control(14)=   "cboSounds"
         Tab(2).Control(15)=   "cmdPlay"
         Tab(2).Control(16)=   "cmdEditor"
         Tab(2).Control(17)=   "chkMinimize"
         Tab(2).ControlCount=   18
         TabCaption(3)   =   "Impresora"
         TabPicture(3)   =   "Main.frx":43D0
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Line(12)"
         Tab(3).Control(1)=   "Line(13)"
         Tab(3).Control(2)=   "lblOutput(4)"
         Tab(3).Control(3)=   "lblOutput(3)"
         Tab(3).Control(4)=   "lblOutput(2)"
         Tab(3).Control(5)=   "lblOutput(1)"
         Tab(3).Control(6)=   "lblOutput(0)"
         Tab(3).Control(7)=   "Label(16)"
         Tab(3).Control(8)=   "lblPort(0)"
         Tab(3).Control(9)=   "lblPort(1)"
         Tab(3).Control(10)=   "lblPort(2)"
         Tab(3).Control(11)=   "lblPort(3)"
         Tab(3).Control(12)=   "lblPort(4)"
         Tab(3).Control(13)=   "lblLeft(1)"
         Tab(3).Control(14)=   "lblTop(1)"
         Tab(3).Control(15)=   "lblBottom(1)"
         Tab(3).Control(16)=   "lblRight(1)"
         Tab(3).Control(17)=   "lblPort(6)"
         Tab(3).Control(18)=   "lblPort(5)"
         Tab(3).Control(19)=   "cboHeight"
         Tab(3).Control(20)=   "cboWidth"
         Tab(3).Control(21)=   "cboPort"
         Tab(3).Control(22)=   "picContainer(1)"
         Tab(3).Control(23)=   "chkFormFeed"
         Tab(3).Control(24)=   "cmdTextTest"
         Tab(3).ControlCount=   25
         TabCaption(4)   =   "Fuente"
         TabPicture(4)   =   "Main.frx":43EC
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lblFont(6)"
         Tab(4).Control(1)=   "lblFont(5)"
         Tab(4).Control(2)=   "lblFont(4)"
         Tab(4).Control(3)=   "lblFont(3)"
         Tab(4).Control(4)=   "lblFont(2)"
         Tab(4).Control(5)=   "lblFont(1)"
         Tab(4).Control(6)=   "lblFont(0)"
         Tab(4).Control(7)=   "cmdFont(6)"
         Tab(4).Control(8)=   "cmdFont(5)"
         Tab(4).Control(9)=   "cmdFont(4)"
         Tab(4).Control(10)=   "cmdFont(3)"
         Tab(4).Control(11)=   "cmdFont(2)"
         Tab(4).Control(12)=   "cmdFont(1)"
         Tab(4).Control(13)=   "cmdFont(0)"
         Tab(4).ControlCount=   14
         Begin VB.CheckBox chkMinimize 
            Caption         =   "Minimizar mientras imprime"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74910
            TabIndex        =   129
            Top             =   2430
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CommandButton cmdEditor 
            Caption         =   "&Cargar archivo en editor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74880
            TabIndex        =   128
            Top             =   3720
            Width           =   1875
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "&Play"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -72600
            TabIndex        =   44
            Top             =   1155
            Width           =   750
         End
         Begin VB.ComboBox cboSounds 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74910
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1155
            Width           =   2220
         End
         Begin VB.ComboBox cboZoom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":4408
            Left            =   -72600
            List            =   "Main.frx":4421
            TabIndex        =   45
            Text            =   "100%"
            Top             =   1920
            Width           =   750
         End
         Begin VB.CommandButton cmdTextTest 
            Caption         =   "Test texto"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -74880
            TabIndex        =   55
            Top             =   3720
            Width           =   1110
         End
         Begin VB.OptionButton optPagePos 
            Caption         =   "Numero de pagina al pie"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74325
            TabIndex        =   36
            Top             =   1140
            Width           =   2670
         End
         Begin VB.OptionButton optPagePos 
            Caption         =   "Numero de pagina en cabezera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74325
            TabIndex        =   35
            Top             =   885
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.CheckBox chkFooter 
            Caption         =   "Pie de pagina"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -73365
            TabIndex        =   32
            Top             =   120
            Value           =   1  'Checked
            Width           =   1530
         End
         Begin VB.CheckBox chkResetPage 
            Caption         =   "&Cambiar numero por archivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74325
            TabIndex        =   34
            Top             =   630
            Width           =   2490
         End
         Begin VB.CheckBox chkTime 
            Caption         =   "Imprimir hora"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74625
            TabIndex        =   38
            Top             =   1650
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "Imprimir fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74625
            TabIndex        =   37
            Top             =   1395
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkHeader 
            Caption         =   "&Cabezera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74895
            TabIndex        =   31
            Top             =   120
            Value           =   1  'Checked
            Width           =   1410
         End
         Begin VB.CheckBox chkPageNumbers 
            Caption         =   "Numero de paginas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74625
            TabIndex        =   33
            Top             =   375
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.ComboBox cboExtention 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":444B
            Left            =   -74910
            List            =   "Main.frx":444D
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   360
            Width           =   3060
         End
         Begin VB.CheckBox chkPlayWaves 
            Caption         =   "Ejecutar eventos de sonido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -74910
            TabIndex        =   42
            Top             =   870
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkFormFeed 
            Alignment       =   1  'Right Justify
            Caption         =   "Forzar nueva pagina"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74550
            TabIndex        =   54
            Tag             =   "  "
            Top             =   2535
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.PictureBox picContainer 
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
            Height          =   1080
            Index           =   1
            Left            =   -74790
            ScaleHeight     =   1080
            ScaleWidth      =   2925
            TabIndex        =   105
            Top             =   345
            Width           =   2925
            Begin VB.CommandButton cmdPickRtf 
               Caption         =   "..."
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2580
               TabIndex        =   50
               Top             =   510
               Width           =   255
            End
            Begin VB.TextBox txtRTFfile 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   270
               TabIndex        =   49
               Text            =   "Listing.rtf"
               Top             =   480
               Width           =   2595
            End
            Begin VB.OptionButton optOutput 
               Caption         =   "Directo al puerto (solo texto)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   0
               TabIndex        =   48
               Top             =   855
               Width           =   2550
            End
            Begin VB.OptionButton optOutput 
               Caption         =   "Archivo Rich-Text & (no imagenes)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   0
               TabIndex        =   47
               Top             =   255
               Width           =   2865
            End
            Begin VB.OptionButton optOutput 
               Caption         =   "Usar &windows (graficos)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Value           =   -1  'True
               Width           =   2565
            End
         End
         Begin VB.ComboBox cboPort 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":444F
            Left            =   -73335
            List            =   "Main.frx":445C
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Tag             =   "  "
            Top             =   1455
            Width           =   735
         End
         Begin VB.ComboBox cboWidth 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":4472
            Left            =   -73335
            List            =   "Main.frx":447C
            TabIndex        =   53
            Tag             =   "  "
            Text            =   "80"
            Top             =   2175
            Width           =   735
         End
         Begin VB.ComboBox cboHeight 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Main.frx":4489
            Left            =   -73335
            List            =   "Main.frx":4493
            TabIndex        =   52
            Tag             =   "  "
            Text            =   "66"
            Top             =   1815
            Width           =   735
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   -72600
            TabIndex        =   56
            Top             =   135
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   -72600
            TabIndex        =   57
            Top             =   525
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   -72600
            TabIndex        =   58
            Top             =   915
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   -72600
            TabIndex        =   59
            Top             =   1305
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   -72600
            TabIndex        =   60
            Top             =   1710
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   -72600
            TabIndex        =   61
            Top             =   2115
            Width           =   750
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Cambiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   -72600
            TabIndex        =   62
            Top             =   2520
            Width           =   750
         End
         Begin VB.CheckBox chkFormIcons 
            Caption         =   "Imprimir iconos de formularios en otra pagina"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            TabIndex        =   14
            Top             =   330
            Width           =   3435
         End
         Begin VB.TextBox txtOwner 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   -74910
            TabIndex        =   40
            Text            =   "Written by Inner Control Business Management"
            Top             =   2580
            Width           =   3060
         End
         Begin VB.TextBox txtOwner 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   -74910
            TabIndex        =   39
            Text            =   "Este fue creado con Project Printer"
            Top             =   2235
            Width           =   3060
         End
         Begin VB.CheckBox chkProcPage 
            Caption         =   "Un procedimiento por pagina"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   21
            Top             =   2130
            Width           =   2790
         End
         Begin VB.CheckBox chkIcon 
            Caption         =   "Ic&ono formulario (picture)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   600
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkSortControls 
            Caption         =   "Ordenar controles por nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   17
            Top             =   1110
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkControlNames 
            Caption         =   "&Nombre de Controles formularios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   855
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkCode 
            Caption         =   "&Codigo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   19
            Top             =   1620
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkSeparator 
            Caption         =   "Separador de &Lineas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   22
            Top             =   2385
            Value           =   1  'Checked
            Width           =   2790
         End
         Begin VB.CheckBox chkProject 
            Caption         =   "&Informacion del Proyecto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   13
            Top             =   120
            Value           =   1  'Checked
            Width           =   3060
         End
         Begin VB.CheckBox chkProcNames 
            Caption         =   "Solo nombres de procedimientos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   20
            Top             =   1875
            Width           =   2790
         End
         Begin VB.CheckBox chkControlPage 
            Caption         =   "Imprimir en paginas distintas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   18
            Top             =   1365
            Width           =   2790
         End
         Begin VB.CheckBox chkIndex 
            Caption         =   "&Indexar paginas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   23
            Top             =   2640
            Width           =   3060
         End
         Begin VB.CheckBox chkSortIndex 
            Caption         =   "Ordenar por nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   375
            TabIndex        =   24
            Top             =   2895
            Width           =   2790
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   21
            X1              =   -74970
            X2              =   -71745
            Y1              =   2685
            Y2              =   2685
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   20
            X1              =   -75000
            X2              =   -71760
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   19
            X1              =   -75000
            X2              =   -71760
            Y1              =   2340
            Y2              =   2340
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   18
            X1              =   -74970
            X2              =   -71745
            Y1              =   2325
            Y2              =   2325
         End
         Begin VB.Label lblSoundFile 
            Caption         =   "Archivo de sonido: n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74865
            TabIndex        =   127
            Top             =   1545
            Width           =   2955
         End
         Begin VB.Label lblZoomDialog 
            Caption         =   "Zoom pagina por defecto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74865
            TabIndex        =   112
            Top             =   1965
            Width           =   2235
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   15
            X1              =   -74985
            X2              =   -71760
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   14
            X1              =   -75015
            X2              =   -71775
            Y1              =   1815
            Y2              =   1815
         End
         Begin VB.Label lblPort 
            Caption         =   "lineas"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   -72675
            TabIndex        =   124
            Top             =   3180
            Width           =   405
         End
         Begin VB.Label lblPort 
            Caption         =   "caracteres"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   -72675
            TabIndex        =   123
            Tag             =   "  "
            Top             =   3435
            Width           =   795
         End
         Begin VB.Label lblRight 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -73020
            TabIndex        =   122
            Tag             =   "  "
            Top             =   3435
            Width           =   330
         End
         Begin VB.Label lblBottom 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -73080
            TabIndex        =   121
            Tag             =   "  "
            Top             =   3180
            Width           =   330
         End
         Begin VB.Label lblTop 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74280
            TabIndex        =   120
            Tag             =   "  "
            Top             =   3180
            Width           =   330
         End
         Begin VB.Label lblLeft 
            Caption         =   "n/a"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   -74340
            TabIndex        =   119
            Tag             =   "  "
            Top             =   3435
            Width           =   330
         End
         Begin VB.Label lblPort 
            AutoSize        =   -1  'True
            Caption         =   "Derecha:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   -73710
            TabIndex        =   118
            Tag             =   "  "
            Top             =   3435
            UseMnemonic     =   0   'False
            Width           =   660
         End
         Begin VB.Label lblPort 
            Caption         =   "Izq:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   -74745
            TabIndex        =   117
            Tag             =   "  "
            Top             =   3435
            UseMnemonic     =   0   'False
            Width           =   375
         End
         Begin VB.Label lblPort 
            Caption         =   "Abajo:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   -73710
            TabIndex        =   116
            Tag             =   "  "
            Top             =   3180
            UseMnemonic     =   0   'False
            Width           =   600
         End
         Begin VB.Label lblPort 
            AutoSize        =   -1  'True
            Caption         =   "Arriba:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   -74775
            TabIndex        =   115
            Tag             =   "  "
            Top             =   3180
            UseMnemonic     =   0   'False
            Width           =   450
         End
         Begin VB.Label lblPort 
            Caption         =   "Margenes texto:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74880
            TabIndex        =   114
            Tag             =   "  "
            Top             =   2925
            UseMnemonic     =   0   'False
            Width           =   2985
         End
         Begin VB.Label lblExtention 
            Caption         =   "Extension para Seleccionar archivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -74865
            TabIndex        =   113
            Top             =   105
            Width           =   2985
         End
         Begin VB.Label Label 
            Caption         =   "Dispositivo de Impresora:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   -74865
            TabIndex        =   111
            Top             =   105
            Width           =   2985
         End
         Begin VB.Label lblOutput 
            Caption         =   "Puerto:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   -74520
            TabIndex        =   110
            Tag             =   "  "
            Top             =   1515
            Width           =   960
         End
         Begin VB.Label lblOutput 
            AutoSize        =   -1  'True
            Caption         =   "Ancho pagina:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   -74520
            TabIndex        =   109
            Tag             =   "  "
            Top             =   2220
            Width           =   1035
         End
         Begin VB.Label lblOutput 
            AutoSize        =   -1  'True
            Caption         =   "Tamaño pagina:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   -74520
            TabIndex        =   108
            Tag             =   "  "
            Top             =   1860
            Width           =   1155
         End
         Begin VB.Label lblOutput 
            Caption         =   "lineas"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   -72450
            TabIndex        =   107
            Tag             =   "  "
            Top             =   1860
            Width           =   405
         End
         Begin VB.Label lblOutput 
            Caption         =   "caracteres"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   -72450
            TabIndex        =   106
            Tag             =   "  "
            Top             =   2220
            Width           =   795
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   13
            Tag             =   "  "
            X1              =   -74985
            X2              =   -71745
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   12
            Tag             =   "  "
            X1              =   -74985
            X2              =   -71760
            Y1              =   2805
            Y2              =   2805
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Procedimientos"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   -74895
            TabIndex        =   104
            Top             =   135
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Codigo"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   -74895
            TabIndex        =   103
            Top             =   525
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Comentarios"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   -74895
            TabIndex        =   102
            Top             =   915
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cabezera"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   -74895
            TabIndex        =   101
            Top             =   1305
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pie de Pagina"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   -74895
            TabIndex        =   100
            Top             =   1710
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Directivas"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   -74895
            TabIndex        =   99
            Top             =   2115
            Width           =   2220
         End
         Begin VB.Label lblFont 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Titulos"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   -74895
            TabIndex        =   98
            Top             =   2520
            UseMnemonic     =   0   'False
            Width           =   2220
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   17
            X1              =   -74985
            X2              =   -71745
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   16
            X1              =   -74985
            X2              =   -71760
            Y1              =   1905
            Y2              =   1905
         End
         Begin VB.Label Label 
            Caption         =   "Texto del pie (ej. Propietario del codigo)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   -74865
            TabIndex        =   97
            Top             =   1980
            Width           =   2985
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   11
            X1              =   -75000
            X2              =   -71760
            Y1              =   780
            Y2              =   780
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   10
            X1              =   -74985
            X2              =   -71760
            Y1              =   765
            Y2              =   765
         End
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&Visualizar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3540
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3900
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "Configurar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6450
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3900
         Visible         =   0   'False
         Width           =   1110
      End
   End
   Begin VB.TextBox txtWrap 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   4815
      MultiLine       =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   450
   End
   Begin RichTextLib.RichTextBox RTBox 
      Height          =   285
      Left            =   3165
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   503
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Main.frx":449F
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   77
      Top             =   5595
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   "Listo"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6105
      Top             =   4785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar Cambios"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   390
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4845
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restaurar Cambios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1845
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4845
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivo_Abrir 
         Caption         =   "|Abrir proyecto o archivo|&Abrir"
      End
      Begin VB.Menu mnuArchivo_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_Visualizar 
         Caption         =   "|Visualizar archivo o modulo seleccionado|&Visualizar"
      End
      Begin VB.Menu mnuArchivo_Configurar 
         Caption         =   "|Configurar impresora|&Configurar impresora"
      End
      Begin VB.Menu mnuArchivo_Imprimir 
         Caption         =   "|Imprimir a impresora o archivo|&Imprimir"
      End
      Begin VB.Menu mnuArchivo_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_Salir 
         Caption         =   "|Salir de la aplicacion|&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edicion"
      Begin VB.Menu mnuEdicion_SelTodo 
         Caption         =   "|Seleccionar todo a imprimir|&Seleccionar Todo"
      End
      Begin VB.Menu mnuEdicion_LimpiarTodo 
         Caption         =   "|Limpiar toda la seleccion|&Limpiar Todo"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOpciones_GCambios 
         Caption         =   "|Guardar configuracion de la aplicacion|&Guardar Cambios"
      End
      Begin VB.Menu mnuOpciones_RCambios 
         Caption         =   "|Restaurar configuracion previa|&Restaurar Cambios"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuAyuda_Indice 
         Caption         =   "|Indice de la ayuda de la aplicacion|&Indice"
      End
      Begin VB.Menu mnuAyuda_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_AcercaDe 
         Caption         =   "|Informacion sobre Copyright y del autor|A&cerca de ..."
      End
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu mnuRecentFile 
         Caption         =   "(Empty)"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuRecentBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCancel 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Comments here belong to the declarations section - these are just here for testing.

Private mGradient As New clsGradient
Private clsXmenu As New CXtremeMenu
Private WithEvents MyHelpCallBack As HelpCallBack
Attribute MyHelpCallBack.VB_VarHelpID = -1

Dim sLoadedFile As String
Dim bBinding As Boolean
Dim bZoomRefresh As Boolean

' These are some comments it the top - belonging to the next
' procedure. A empty line above marks the seperation.
' So long no statements are below these comments, these will belong to the procedure.

Private Sub Form_Load()
   On Error Resume Next

   bBinding = True

   ' Set the controls...
   cboPort.ListIndex = 0

   cboExtention.AddItem "Archivos VB (*.vbp;*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob)"
   cboExtention.AddItem "Proyectos (*.vbp)"
   cboExtention.AddItem "Formularios (*.frm)"
   cboExtention.AddItem "Modulos (*.bas)"
   cboExtention.AddItem "Clases (*.cls)"
   cboExtention.AddItem "Controles de Usuario (*.ctl)"
   cboExtention.AddItem "Paginas de Propiedades (*.pag)"
   cboExtention.AddItem "Documentos de Usuario (*.dob)"
   cboExtention.AddItem "Todos los archivos (*.*)"
   cboExtention.ListIndex = 0

   cboSounds.AddItem "Accesando archivo"
   cboSounds.AddItem "Analizando"
   cboSounds.AddItem "Error"
   cboSounds.AddItem "Salir aplicacion"
   cboSounds.AddItem "Ok"
   cboSounds.AddItem "Listo"
   cboSounds.AddItem "Disculpa, ..."
   cboSounds.AddItem "Ocupado"
   cboSounds.AddItem "Inicio"
   cboSounds.AddItem "Gracias"
   cboSounds.ListIndex = 0

   chkPlayWaves.Value = IIf(GetIniString(sIniFile, "Options", "WaveSounds", "1") = "1", vbChecked, vbUnchecked)
   chkMinimize.Value = IIf(GetIniString(sIniFile, "Options", "Minimize", "1") = "1", vbChecked, vbUnchecked)

   ' Read INI file and set the recent menu file list control array appropriately.
   GetRecentFiles

   RestoreFromINI
   ButtonsState
   ShowPrinterInfo

   bBinding = False

    Set MyHelpCallBack = New HelpCallBack
    
    Call clsXmenu.Install(hWnd, MyHelpCallBack, Me.imlMain)
    Call clsXmenu.FontName(hWnd, "Tahoma")
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision, picDraw)
    
    picDraw.Refresh
    
   CentreForm Me
   Me.Show
End Sub



Private Sub Form_Unload(Cancel As Integer)
   If cmdSave.Enabled Then
      Select Case MsgBox("Cambios a la configuracion. Guardar cambios realizados?", vbYesNoCancel + vbQuestion + vbDefaultButton2)
      Case vbYes
         cmdSave_Click
      Case vbCancel
         Cancel = True
         Exit Sub
      End Select
   End If

   On Error Resume Next
   Unload frmPrint
   Unload frmPreview

   MakeSound WAVE_EXIT, True
   End                  ' Just in case anything is still running...
End Sub


Private Sub cmdHelp_Click()
   On Error Resume Next
   MousePointer = vbHourglass

   Load frmViewFile
   frmViewFile.ShowHelpFile

   MousePointer = vbDefault

   frmViewFile.Show
End Sub

Private Sub lblPrinter_DblClick()
   SSTab.Tab = 1
   TabOptions.Tab = 3
   If optOutput(1) Then
      optOutput(1).SetFocus
   ElseIf optOutput(2) Then
      optOutput(2).SetFocus
   Else
      optOutput(0).SetFocus
   End If
End Sub



Private Sub lblViewSize_DblClick()
   SSTab.Tab = 1
   TabOptions.Tab = 2
   cboZoom.SetFocus
End Sub

' --------------------------------------------------------
Private Sub cmdSave_Click()
   Dim i As Integer
   Dim sText As String
   Dim aFonts As Variant

   On Error GoTo SaveError
   AddIniString sIniFile, "Options", "Header", IIf(chkHeader = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "Footer", IIf(chkFooter = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "PageNumbers", IIf(chkPageNumbers = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ResetPage", IIf(chkResetPage = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "DateStamp", IIf(chkDate = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "TimeStamp", IIf(chkTime = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "Index", IIf(chkIndex = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "SortIndex", IIf(chkSortIndex = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ProjectInfo", IIf(chkProject = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "AllIcons", IIf(chkFormIcons = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "FormIcon", IIf(chkIcon = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ControlNames", IIf(chkControlNames = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "SortControls", IIf(chkSortControls = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ControlPage", IIf(chkControlPage = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "Code", IIf(chkCode = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "NamesOnly", IIf(chkProcNames = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "ProcPerPage", IIf(chkProcPage = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "SubSeparator", IIf(chkSeparator = vbChecked, "1", "0")
   AddIniString sIniFile, "Options", "PagePos", IIf(optPagePos(1), "Footer", "Header")
   AddIniString sIniFile, "Options", "PreviewZoom", cboZoom.Text
   AddIniString sIniFile, "Options", "Extention", cboExtention.ListIndex

   AddIniString sIniFile, "Print", "Device", IIf(optOutput(1), "File", IIf(optOutput(2), "Port", "Driver"))
   AddIniString sIniFile, "Print", "RTFFile", txtRTFfile
   AddIniString sIniFile, "Print", "Port", cboPort
   AddIniString sIniFile, "Print", "Width", cboWidth
   AddIniString sIniFile, "Print", "Height", cboHeight
   AddIniString sIniFile, "Print", "FormFeed", IIf(chkFormFeed = vbChecked, "1", "0")

   AddIniString sIniFile, "Margins", "Top", lblTop(0)
   AddIniString sIniFile, "Margins", "Bottom", lblBottom(0)
   AddIniString sIniFile, "Margins", "Left", lblLeft(0)
   AddIniString sIniFile, "Margins", "Right", lblRight(0)

   AddIniString sIniFile, "Footer", "Line1", txtOwner(0).Text
   AddIniString sIniFile, "Footer", "Line2", txtOwner(1).Text

   aFonts = Array("Procs", "Code", "Comments", "Header", "Footer", "Directive", "Titles")

   For i = 0 To 6
      sText = "Font" & aFonts(i)

      AddIniString sIniFile, sText, "Font", lblFont(i).FontName
      AddIniString sIniFile, sText, "Size", lblFont(i).FONTSIZE
      AddIniString sIniFile, sText, "Color", "&H" & Right("00000000" & Hex$(lblFont(i).ForeColor), 8)
      AddIniString sIniFile, sText, "Bold", IIf(lblFont(i).FontBold, "1", "0")
      AddIniString sIniFile, sText, "Italic", IIf(lblFont(i).FontItalic, "1", "0")
      AddIniString sIniFile, sText, "Strikethru", IIf(lblFont(i).FontStrikethru, "1", "0")
      AddIniString sIniFile, sText, "Underline", IIf(lblFont(i).FontUnderline, "1", "0")
   Next

   SetEnabled cmdSave, False
   SetEnabled cmdRestore, False

   MakeSound WAVE_OK, True
   Exit Sub
SaveError:
   MsgBox "Problema guardando configuracion." & vbCrLf & _
          "Error reportado #" & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub cmdRestore_Click()
   If MsgBox("Cambios realizados se perderan! Esta seguro de restaurar configuracion previa?", vbYesNo + vbQuestion) = vbYes Then
      RestoreFromINI
      SetMargins
   End If
End Sub

Private Sub RestoreFromINI()
   Dim sText As String
   Dim nNumber As Integer, i As Integer
   Dim aFonts As Variant
   On Error Resume Next

   chkHeader = IIf(GetIniString(sIniFile, "Options", "Header", IIf(chkHeader = vbChecked, "1", "0")) = "1", 1, 0)
   chkFooter = IIf(GetIniString(sIniFile, "Options", "Footer", IIf(chkFooter = vbChecked, "1", "0")) = "1", 1, 0)
   chkPageNumbers = IIf(GetIniString(sIniFile, "Options", "PageNumbers", IIf(chkPageNumbers = vbChecked, "1", "0")) = "1", 1, 0)
   chkResetPage = IIf(GetIniString(sIniFile, "Options", "resetPage", IIf(chkResetPage = vbChecked, "1", "0")) = "1", 1, 0)
   chkDate = IIf(GetIniString(sIniFile, "Options", "DateStamp", IIf(chkDate = vbChecked, "1", "0")) = "1", 1, 0)
   chkTime = IIf(GetIniString(sIniFile, "Options", "TimeStamp", IIf(chkTime = vbChecked, "1", "0")) = "1", 1, 0)
   chkIndex = IIf(GetIniString(sIniFile, "Options", "Index", IIf(chkIndex = vbChecked, "1", "0")) = "1", 1, 0)
   chkSortIndex = IIf(GetIniString(sIniFile, "Options", "SortIndex", IIf(chkSortIndex = vbChecked, "1", "0")) = "1", 1, 0)
   chkProject = IIf(GetIniString(sIniFile, "Options", "ProjectInfo", IIf(chkProject = vbChecked, "1", "0")) = "1", 1, 0)
   chkFormIcons = IIf(GetIniString(sIniFile, "Options", "AllIcons", IIf(chkFormIcons = vbChecked, "1", "0")) = "1", 1, 0)
   chkIcon = IIf(GetIniString(sIniFile, "Options", "FormIcon", IIf(chkIcon = vbChecked, "1", "0")) = "1", 1, 0)
   chkControlNames = IIf(GetIniString(sIniFile, "Options", "ControlNames", IIf(chkControlNames = vbChecked, "1", "0")) = "1", 1, 0)
   chkSortControls = IIf(GetIniString(sIniFile, "Options", "SortControls", IIf(chkSortControls = vbChecked, "1", "0")) = "1", 1, 0)
   chkControlPage = IIf(GetIniString(sIniFile, "Options", "ControlPage", IIf(chkControlPage = vbChecked, "1", "0")) = "1", 1, 0)
   chkCode = IIf(GetIniString(sIniFile, "Options", "Code", IIf(chkCode = vbChecked, "1", "0")) = "1", 1, 0)
   chkProcNames = IIf(GetIniString(sIniFile, "Options", "NamesOnly", IIf(chkProcNames = vbChecked, "1", "0")) = "1", 1, 0)
   chkProcPage = IIf(GetIniString(sIniFile, "Options", "ProcPerPage", IIf(chkProcPage = vbChecked, "1", "0")) = "1", 1, 0)
   chkSeparator = IIf(GetIniString(sIniFile, "Options", "SubSeparator", IIf(chkSeparator = vbChecked, "1", "0")) = "1", 1, 0)
   If UCase$(Left(GetIniString(sIniFile, "Options", "PagePos", "Header"), 1)) = "F" Then
      optPagePos(1) = True
   Else
      optPagePos(0) = True
   End If
   nNumber = Int(Val(GetIniString(sIniFile, "Options", "PreviewZoom", "100%")))
   If nNumber <= 0 Then
      cboZoom.Text = "Fit"
   Else
      cboZoom.Text = Format(nNumber, "##0") & "%"
   End If
   nNumber = Val(GetIniString(sIniFile, "Options", "Extention", cboExtention.ListIndex))
   cboExtention.ListIndex = IIf(nNumber < 0 Or nNumber > 8, 0, nNumber)

   sText = UCase$(Left(GetIniString(sIniFile, "Print", "Device", "Driver"), 1))
   If sText = "F" Then
      optOutput(1) = True
   ElseIf sText = "P" Then
      optOutput(2) = True
   Else
      optOutput(0) = True
   End If
   txtRTFfile = GetIniString(sIniFile, "Print", "RTFFile", txtRTFfile)
   cboPort = GetIniString(sIniFile, "Print", "Port", cboPort)
   cboWidth = GetIniString(sIniFile, "Print", "Width", cboWidth)
   cboHeight = GetIniString(sIniFile, "Print", "Height", cboHeight)
   chkFormFeed = IIf(GetIniString(sIniFile, "Print", "FormFeed", IIf(chkFormFeed = vbChecked, "1", "0")) = "1", 1, 0)

   lblTop(0) = GetIniString(sIniFile, "Margins", "Top", lblTop(0))
   lblBottom(0) = GetIniString(sIniFile, "Margins", "Bottom", lblBottom(0))
   lblLeft(0) = GetIniString(sIniFile, "Margins", "Left", lblLeft(0))
   lblRight(0) = GetIniString(sIniFile, "Margins", "Right", lblRight(0))
   AssignSliderValues

   txtOwner(0).Text = GetIniString(sIniFile, "Footer", "Line1", "Created with 'VB.Print!' - VB Source code printing utility")
   txtOwner(1).Text = GetIniString(sIniFile, "Footer", "Line2", "'VB.Print!' is written by Inner Control Business Management (Australia)")

   aFonts = Array("Procs", "Code", "Comments", "Header", "Footer", "Directive", "Titles")

   For i = 0 To 6
      sText = "Font" & aFonts(i)

      lblFont(i).FontName = GetIniString(sIniFile, sText, "Font", lblFont(i).FontName)
      lblFont(i).FONTSIZE = Val(GetIniString(sIniFile, sText, "Size", lblFont(i).FONTSIZE))
      lblFont(i).ForeColor = Abs(Val(GetIniString(sIniFile, sText, "Color", "&H" & Hex(lblFont(i).ForeColor))))
      lblFont(i).FontBold = (GetIniString(sIniFile, sText, "Bold", IIf(lblFont(i).FontBold, "1", "0")) = "1")
      lblFont(i).FontItalic = (GetIniString(sIniFile, sText, "Italic", IIf(lblFont(i).FontItalic, "1", "0")) = "1")
      lblFont(i).FontStrikethru = (GetIniString(sIniFile, sText, "Strikethru", IIf(lblFont(i).FontStrikethru, "1", "0")) = "1")
      lblFont(i).FontUnderline = (GetIniString(sIniFile, sText, "Underline", IIf(lblFont(i).FontUnderline, "1", "0")) = "1")
   Next

   SetEnabled cmdSave, False
   SetEnabled cmdRestore, False

End Sub

Private Sub chkPlayWaves_Click()
   AddIniString sIniFile, "Options", "WaveSounds", IIf(chkPlayWaves = vbChecked, "1", "0")
End Sub

Private Sub chkMinimize_Click()
   AddIniString sIniFile, "Options", "Minimize", IIf(chkMinimize = vbChecked, "1", "0")
End Sub

Private Sub cmdEditor_Click()
   On Error GoTo EditorCancelled

   CommonDialog.DialogTitle = "Open"
   CommonDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|Text document (*.txt)|*.txt|All files (*.*)|*.*"
   CommonDialog.FilterIndex = 1

   CommonDialog.CancelError = True
   CommonDialog.Flags = cdlOFNHideReadOnly
   CommonDialog.FileName = ""

   CommonDialog.ShowOpen

   CommonDialog.CancelError = False

   MousePointer = vbHourglass

   Load frmViewFile
   frmViewFile.SetFileName CommonDialog.FileName
   frmViewFile.InitView

   MousePointer = vbDefault

   frmViewFile.Show

EditorCancelled:
   CommonDialog.CancelError = False
End Sub

Private Sub cmdView_Click()
   On Error Resume Next

   If Outline.ListCount < 1 Or Outline.ListIndex < 0 Then Exit Sub

   Dim sText As String
   Dim i As Integer, n As Integer

   MousePointer = vbHourglass

   Load frmViewFile

   Select Case ItemRef(Outline.ListIndex).ProcPoint
   Case Is = 0
      sText = ""
      For i = 1 To Mdl(ItemRef(Outline.ListIndex).FilePoint).CtrlCount

         If Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Elements > 1 Then
            sText = sText & Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Name, 20) & " " & _
                            Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Type, 20) & " " & _
                            Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Library, 20) & " " & _
                            "Elementos: " & Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Elements
         Else
            sText = sText & Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Name, 20) & " " & _
                            Pad(Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Type, 20) & " " & _
                            Mdl(ItemRef(Outline.ListIndex).FilePoint).Control(i).Library
         End If
         sText = sText & vbCrLf
      Next
      sText = sText & vbCrLf & _
              "   Total nombre controles: " & Mdl(ItemRef(Outline.ListIndex).FilePoint).CtrlCount & vbCrLf & _
              "Total elementos controles: " & Mdl(ItemRef(Outline.ListIndex).FilePoint).CtrlElements

      frmViewFile.SetText "Controles Formulario", sText
    
   Case Is > 0
      Dim sUpper As String
      sText = ""
      n = 1       ' Remove empty line(s) in top
      For i = 1 To Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Lines
         If Not EmptyString(Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Code(i)) Then
            n = i
            Exit For
         End If
      Next
      frmViewFile.Caption = Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).IndexName & " - View"
      frmViewFile.SetFont FONT_CODE

      For i = n To Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Lines
         
         sText = Mdl(ItemRef(Outline.ListIndex).FilePoint).Proc(ItemRef(Outline.ListIndex).ProcPoint).Code(i) & vbCrLf
         sUpper = UCase$(Trim$(sText))

         If MatchString(sUpper, "'") Then                      ' Comments
            frmViewFile.SetFont FONT_COMMENTS
            frmViewFile.SetLine sText
            frmViewFile.SetFont FONT_CODE

         ElseIf MatchString(sUpper, "#") Then                  ' Compiler directive
            frmViewFile.SetFont FONT_DIRECTIVE
            frmViewFile.SetLine sText
            frmViewFile.SetFont FONT_CODE

         ElseIf IsProcedure(sUpper) Then                       ' Only happens once (I hope) - and not in declaration section
            frmViewFile.SetFont FONT_PROCS
            frmViewFile.SetLine sText
            frmViewFile.SetFont FONT_CODE
         Else                                                  ' Just some code or space
            frmViewFile.SetLine sText
         End If
      Next
 
   Case Else
      frmViewFile.SetFileName Mdl(ItemRef(Outline.ListIndex).FilePoint).PathFile
   End Select

   frmViewFile.InitView

   MousePointer = vbDefault

   frmViewFile.Show
End Sub

' --------------------------------------------------------

Private Sub cboSounds_Change()
   cboSounds_Click
End Sub

Private Sub cboSounds_Click()
   lblSoundFile.Caption = "Wave file: " & GetSoundFileName(cboSounds.ListIndex)
End Sub

Private Sub cmdPlay_Click()
   If bBinding Then Exit Sub
   If cboSounds.ListIndex < 0 Then Exit Sub
   MakeSound cboSounds.ListIndex, False, True
End Sub

' --------------------------------------------------------

' Test layout of text printer
Private Sub cmdTextTest_Click()
   On Error GoTo TestPrintError

   SetEnabled cmdTextTest, False

   Me.MousePointer = vbHourglass
   DoEvents
   frmMain.Enabled = False
   On Error GoTo 0

   MakeSound WAVE_STANDBY

   TestTextPrint

TestPrintError:
   frmMain.Enabled = True
   SetEnabled cmdTextTest, True
   Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrintSetup_Click()
   CommonDialog.CancelError = False
   CommonDialog.Flags = cdlPDPrintSetup
   CommonDialog.ShowPrinter

   DoEvents
   ShowPrinterInfo
End Sub

Public Sub SetRTFfile(sFIle As String)
   bBinding = True
   frmMain.txtRTFfile = sFIle
   bBinding = False
End Sub

Private Sub cmdPrint_Click()
   Dim nFormState As Integer
   nFormState = Me.WindowState

   On Error GoTo PrintDialogCancelled

   SetEnabled cmdPrint, False
   tbrMain.Buttons("cmdImprimir").Enabled = False

   If optOutput(1) Then                ' RTF
      Dim sRTFFile As String
      Me.MousePointer = vbHourglass
      DoEvents
      frmMain.Enabled = False
      On Error GoTo 0

      If frmMain.chkPreview <> vbChecked Then
         sRTFFile = frmMain.txtRTFfile
         If Not FileOverwriteDialog(sRTFFile, CommonDialog, "RTF files (*.rtf)|*.rtf|All files (*.*)|*.*", ".rtf") Then
            GoTo PrintDialogCancelled
         End If
         SetRTFfile sRTFFile
      End If
   
   ElseIf optOutput(2) Then            ' Port
      Me.MousePointer = vbHourglass
      DoEvents
      frmMain.Enabled = False
      On Error GoTo 0

   Else
      If chkPreview <> vbChecked Then
         CommonDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection Or cdlPDUseDevModeCopies Or cdlPDPageNums 'Or cdlPDPageNums 'Or cdlPDNoPageNums
'         CommonDialog.FromPage = txtFromPage
'         CommonDialog.ToPage = txtToPage
'         CommonDialog.Min = 1
'         CommonDialog.Max = 1

         CommonDialog.FromPage = 1
         CommonDialog.ToPage = 1
         CommonDialog.Min = 1
         CommonDialog.Max = 1

         CommonDialog.CancelError = True

         CommonDialog.ShowPrinter

         CommonDialog.CancelError = False
      End If

      Me.MousePointer = vbHourglass
      DoEvents
      If chkPreview <> vbChecked Then ShowPrinterInfo
      frmMain.Enabled = False
      On Error GoTo 0
   End If

   MakeSound WAVE_STANDBY

   If GetIniString(sIniFile, "Options", "Minimize", "1") = "1" Then
      WindowState = vbMinimized
   End If

   ' This is it - The whole application turns around this routine.
   PrintControl

PrintDialogCancelled:
   lstNames(0).Clear
   lstNames(1).Clear
   picImage.Picture = LoadPicture()
   frmMain.Enabled = True
   SetEnabled cmdPrint, True
   tbrMain.Buttons("cmdImprimir").Enabled = True
   CommonDialog.CancelError = False
   Me.MousePointer = vbDefault

   If WindowState = vbMinimized Then
      DoEvents
      Do While Page.Show: DoEvents: Loop
      If WindowState = vbMinimized Then WindowState = nFormState
   End If
End Sub

Private Sub ShowPrinterInfo(Optional bRefreshSample)
   Me.MousePointer = vbHourglass

   If optOutput(2) Then
      lblPrinter = "Archivo plano o impresora " & cboPort
      lblOrient = "n/a"
      lblSize = Format(cboWidth, "###0") & " (caracteres) por " & Format(cboHeight, "###0") & " (lineas)"

   Else
      If optOutput(1) Then
         lblPrinter = "Archivo RTF " & txtRTFfile
      Else
         lblPrinter = Printer.DeviceName & " en " & Printer.Port
      End If

      Select Case Printer.Orientation
      Case vbPRORPortrait
         lblOrient = "Portrait"
      Case vbPRORLandscape
         lblOrient = "Landscape"
      Case Else
         lblOrient = "n/a"
      End Select

      'Printer.ScaleLeft, Printer.ScaleTop
      Printer.ScaleMode = vbMillimeters
      If optOutput(1) Then
         lblSize = Format(Printer.ScaleWidth - (Val(lblLeft(0)) + Val(lblRight(0))), "###0") & " mm por " & Format(Printer.ScaleHeight - (Val(lblTop(0)) + Val(lblBottom(0))), "###0") & " mm"
      Else
         lblSize = Format(Printer.ScaleWidth, "###0") & " mm por " & Format(Printer.ScaleHeight, "###0") & " mm"
      End If
   End If

   If IsMissing(bRefreshSample) Then bRefreshSample = True

   If bRefreshSample Then
      ' Don't change the execution order !!
      SetSampleOrientation
      SetSliders
      If Page.Show Then
         PaintSamplePage
      Else
         PrintSamplePage
      End If
   End If

   Me.MousePointer = vbDefault
End Sub

Private Sub mnuArchivo_Abrir_Click()
    cmdPickFile_Click
End Sub

Private Sub mnuArchivo_Configurar_Click()
    cmdPrintSetup_Click
End Sub

Private Sub mnuArchivo_Imprimir_Click()
    cmdPrint_Click
End Sub

Private Sub mnuArchivo_Salir_Click()
    Unload Me
End Sub

Private Sub mnuArchivo_Visualizar_Click()
    cmdView_Click
End Sub

Private Sub mnuAyuda_AcercaDe_Click()
    frmAcerca.Show vbModal
End Sub

Private Sub mnuEdicion_LimpiarTodo_Click()
    cmdClear_Click
End Sub

Private Sub mnuEdicion_SelTodo_Click()
    cmdSelectAll_Click
End Sub

Private Sub mnuOpciones_GCambios_Click()
    cmdSave_Click
End Sub

Private Sub mnuOpciones_RCambios_Click()
    cmdRestore_Click
End Sub

Private Sub MyHelpCallBack_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    
    StatusBar.SimpleText = MenuHelp
    
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "cmdAbrir"
            cmdPickFile_Click
        Case "cmdSelTodo"
            cmdSelectAll_Click
        Case "cmdLimTodo"
            cmdClear_Click
        Case "cmdPreview"
            mnuArchivo_Visualizar_Click
        Case "cmdImprimir"
            mnuArchivo_Imprimir_Click
        Case "cmdSalir"
            Unload Me
    End Select
    
End Sub

' --------------------------------------------------------

Private Sub txtProject_Change()
   If Len(Trim(txtProject)) = 0 Then
      If Outline.ListCount > 0 Then
         ClearOutline
         ButtonsState
      End If
   Else
      If Not FileExist(txtProject) Then
         If Outline.ListCount > 0 Then
            ClearOutline
            ButtonsState
         End If
      End If
   End If
End Sub

Private Sub txtProject_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Len(Trim(txtProject)) > 0 Then
         GetProjectDetails
      End If
   End If
End Sub

Private Sub cmdPrevFile_Click()
'   Dim sText As String
'   sText = GetIniString(sIniFile, "Options", "LastFile", "")
'   If Len(Trim(sText)) > 0 Then
'      txtProject = sText
'      GetProjectDetails
'   End If
   PopupMenu mnuFiles
End Sub

Private Sub mnuRecentFile_Click(Index As Integer)
   If Len(Trim$(mnuRecentFile(Index).Caption)) > 0 Then
      If txtProject <> mnuRecentFile(Index).Caption Then
         txtProject = mnuRecentFile(Index).Caption
         GetProjectDetails
      End If
   End If
End Sub

Private Sub cmdPickFile_Click()
   On Error GoTo PickFileCancelled

   CommonDialog.DialogTitle = "Seleccionar archivo"
   CommonDialog.Filter = "Archivos VB (*.vbp;*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob)|*.vbp;*.frm;*.bas;*.cls;*.ctl;*.pag;*.dob|" & _
                         "Proyectos (*.vbp)|*.vbp|" & _
                         "Formularios (*.frm)|*.frm|" & _
                         "Modulos (*.bas)|*.bas|" & _
                         "Modulos de Clase (*.cls)|*.cls|" & _
                         "Controles de Usuario (*.ctl)|*.ctl|" & _
                         "Paginas de propiedades (*.pag)|*.pag|" & _
                         "Documentos de usuario (*.dob)|*.dob|" & _
                         "Todos los archivos (*.*)|*.*"

   CommonDialog.FilterIndex = cboExtention.ListIndex + 1

   CommonDialog.CancelError = True
   CommonDialog.Flags = cdlOFNHideReadOnly
   CommonDialog.FileName = txtProject

   CommonDialog.ShowOpen

   CommonDialog.CancelError = False

   txtProject = CommonDialog.FileName

   GetProjectDetails

PickFileCancelled:
   CommonDialog.CancelError = False
End Sub



' --- Preferences tab area -----------------------------------------------------

Private Sub SSTab_Click(PreviousTab As Integer)
   SetVisible cmdSave, (SSTab.Tab = 1)
   SetVisible cmdRestore, (SSTab.Tab = 1)
End Sub

Private Sub chkHeader_Click()
   SetEnabled chkPageNumbers, (chkHeader = vbChecked Or chkFooter = vbChecked)
   SetEnabled chkResetPage, (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(0), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(1), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkFooter_Click()
   SetEnabled chkPageNumbers, (chkHeader = vbChecked Or chkFooter = vbChecked)
   SetEnabled chkResetPage, (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(0), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled optPagePos(1), (chkPageNumbers.Enabled And (chkPageNumbers = vbChecked))
   SetEnabled chkDate, (chkFooter = vbChecked)
   SetEnabled chkTime, (chkFooter = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkPageNumbers_Click()
   SetEnabled chkResetPage, (chkPageNumbers = vbChecked)
   SetEnabled optPagePos(0), (chkPageNumbers = vbChecked)
   SetEnabled optPagePos(1), (chkPageNumbers = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkResetPage_Click()
   EnabledStorage
End Sub

Private Sub optPagePos_Click(Index As Integer)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkDate_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkTime_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkIndex_Click()
   SetEnabled chkSortIndex, (chkIndex = vbChecked)
   EnabledStorage
End Sub

Private Sub chkSortIndex_Click()
   EnabledStorage
End Sub

Private Sub chkProject_Click()
   ButtonsState
   EnabledStorage
End Sub

Private Sub chkFormIcons_Click()
   EnabledStorage
End Sub

Private Sub chkIcon_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkControlNames_Click()
   SetEnabled chkSortControls, (chkControlNames = vbChecked)
   SetEnabled chkControlPage, (chkControlNames = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkSortControls_Click()
   EnabledStorage
End Sub

Private Sub chkControlPage_Click()
   EnabledStorage
End Sub

Private Sub chkCode_Click()
   SetEnabled chkProcNames, (chkCode = vbChecked)
   SetEnabled chkProcPage, (chkCode = vbChecked)
   SetEnabled chkSeparator, (chkCode = vbChecked)
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkProcNames_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkProcPage_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub chkSeparator_Click()
   EnabledStorage
   PaintSamplePage
End Sub

Private Sub cboZoom_Click()
   cboZoom_Change
End Sub

Private Sub cboZoom_Change()
   If bZoomRefresh Then Exit Sub

   Dim nFactor As Single
   nFactor = Int(Val(cboZoom.Text))
   If nFactor <= 0 Then
      lblViewSize = "(Zoom to Fit)"
   Else
      nFactor = nFactor / 100
      lblViewSize = "(Zoom " & Format(nFactor * 100, "##0") & "%)"
   End If
   EnabledStorage
End Sub

Private Sub cboZoom_LostFocus()
   Dim nFactor As Single
   nFactor = Int(Val(cboZoom.Text))

   bZoomRefresh = True
   If nFactor <= 0 Then
      cboZoom.Text = "Fit"
   Else
      nFactor = nFactor / 100
      cboZoom.Text = Format(nFactor * 100, "##0") & "%"
   End If
   bZoomRefresh = False
End Sub

Private Sub txtOwner_LostFocus(Index As Integer)
   PaintSamplePage
End Sub

Private Sub txtOwner_Change(Index As Integer)
   EnabledStorage
End Sub

Private Sub optOutput_Click(Index As Integer)
   If Index = 2 Then
      ' Port
      SetEnabled txtRTFfile, False
      SetEnabled cmdPickRtf, False

      SetEnabled chkFormIcons, False
      SetEnabled chkIcon, False

      SetEnabled lblOutput(0), True
      SetEnabled lblOutput(1), True
      SetEnabled lblOutput(2), True
      SetEnabled lblOutput(3), True
      SetEnabled lblOutput(4), True
      SetEnabled cboPort, True
      SetEnabled cboWidth, True
      SetEnabled cboHeight, True
      SetEnabled chkFormFeed, True

      SetEnabled lblPort(0), True
      SetEnabled lblPort(1), True
      SetEnabled lblPort(2), True
      SetEnabled lblPort(3), True
      SetEnabled lblPort(4), True
      SetEnabled lblPort(5), True
      SetEnabled lblPort(6), True

      SetEnabled lblTop(1), True
      SetEnabled lblBottom(1), True
      SetEnabled lblLeft(1), True
      SetEnabled lblRight(1), True

      SetEnabled cmdTextTest, True

   Else
      If Index = 1 Then
         ' RTF
         SetEnabled txtRTFfile, True
         SetEnabled cmdPickRtf, True

         SetEnabled chkFormIcons, False
         SetEnabled chkIcon, False

      Else
         ' Driver
         SetEnabled txtRTFfile, False
         SetEnabled cmdPickRtf, False

         SetEnabled chkFormIcons, True
         SetEnabled chkIcon, True
      End If

      SetEnabled lblOutput(0), False
      SetEnabled lblOutput(1), False
      SetEnabled lblOutput(2), False
      SetEnabled lblOutput(3), False
      SetEnabled lblOutput(4), False
      SetEnabled cboPort, False
      SetEnabled cboWidth, False
      SetEnabled cboHeight, False
      SetEnabled chkFormFeed, False

      SetEnabled lblPort(0), False
      SetEnabled lblPort(1), False
      SetEnabled lblPort(2), False
      SetEnabled lblPort(3), False
      SetEnabled lblPort(4), False
      SetEnabled lblPort(5), False
      SetEnabled lblPort(6), False

      SetEnabled lblTop(1), False
      SetEnabled lblBottom(1), False
      SetEnabled lblLeft(1), False
      SetEnabled lblRight(1), False

      SetEnabled cmdTextTest, False
   End If

   If bBinding Then Exit Sub
   ShowPrinterInfo False
   ButtonsState
   EnabledStorage
End Sub

Private Sub cmdPickRtf_Click()
   On Error GoTo PickRTFCancelled

   CommonDialog.DialogTitle = "Save RTF file as ..."
   CommonDialog.Filter = "RTF files (*.rtf)|*.rtf|All files (*.*)|*.*"
   CommonDialog.FilterIndex = 1
   CommonDialog.DefaultExt = ".rtf"

   CommonDialog.CancelError = True
   CommonDialog.Flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist

   CommonDialog.FileName = txtRTFfile

   CommonDialog.ShowSave

   txtRTFfile = CommonDialog.FileName

PickRTFCancelled:
   CommonDialog.CancelError = False
End Sub

Private Sub txtRTFfile_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboPort_Click()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboPort_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboHeight_KeyPress(KeyAscii As Integer)
   KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub cboHeight_Click()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboHeight_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboWidth_KeyPress(KeyAscii As Integer)
   KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub cboWidth_Click()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub cboWidth_Change()
   If bBinding Then Exit Sub
   ShowPrinterInfo False
   EnabledStorage
End Sub

Private Sub chkFormFeed_Click()
   EnabledStorage
End Sub

Private Sub cboExtention_Click()
   EnabledStorage
End Sub

Private Sub sldLeft_Change()
   EnabledStorage
   SetMargins
End Sub

Private Sub sldRight_Change()
   EnabledStorage
   SetMargins
End Sub

Private Sub sldTop_Change()
   EnabledStorage
   SetMargins
End Sub

Private Sub sldBottom_Change()
   EnabledStorage
   SetMargins
End Sub

' Only use third of width and length of page for margins
Private Sub SetSliders()
   Dim nMax As Integer, nValue As Integer

   Printer.ScaleMode = vbMillimeters

   nMax = Printer.ScaleWidth / 3
   sldLeft.Max = nMax

   nValue = sldRight.Max - sldRight.Value
   sldRight.Max = nMax
   sldRight.Value = nMax - nValue

   nMax = Printer.ScaleHeight / 3
   sldTop.Max = nMax

   nValue = sldBottom.Max - sldBottom.Value
   sldBottom.Max = nMax
   sldBottom.Value = nMax - nValue
End Sub

Private Sub AssignSliderValues()
   sldLeft.Value = lblLeft(0)
   sldRight.Value = sldRight.Max - lblRight(0)
   sldTop.Value = lblTop(0)
   sldBottom.Value = sldBottom.Max - lblBottom(0)

   SetTextMargins
End Sub

' Margins are in millimeters
Private Sub SetMargins()
   If bBinding Then Exit Sub
   lblLeft(0) = sldLeft.Value
   lblRight(0) = sldRight.Max - sldRight.Value
   lblTop(0) = sldTop.Value
   lblBottom(0) = sldBottom.Max - sldBottom.Value

   If optOutput(1) Then
      lblSize = Format(Printer.ScaleWidth - (Val(lblLeft(0)) + Val(lblRight(0))), "###0") & " mm by " & Format(Printer.ScaleHeight - (Val(lblTop(0)) + Val(lblBottom(0))), "###0") & " mm"
   End If

   SetTextMargins

   PaintSamplePage
End Sub

' The source margins are millimeters. Convert them to characters
Private Sub SetTextMargins()
   Dim nCharFactor As Double, nLineFactor As Double
   nCharFactor = 1200 / 567
   nLineFactor = 2400 / 567
   lblTop(1) = RoundToInt(lblTop(0) / nLineFactor)
   lblBottom(1) = RoundToInt(lblBottom(0) / nLineFactor)
   lblLeft(1) = RoundToInt(lblLeft(0) / nCharFactor)
   lblRight(1) = RoundToInt(lblRight(0) / nCharFactor)
End Sub

' Only accepts positive numbers
Private Function RoundToInt(ByVal nValue As Double) As Integer
   If nValue < 0 Then
      RoundToInt = CInt(nValue)
   Else
      Dim nFraction As Double
      nFraction = nValue - Int(nValue)
      If nFraction < 0.5 Then
         RoundToInt = Int(nValue)
      Else
         RoundToInt = Int(nValue) + 1
      End If
   End If
End Function

Private Sub cmdFont_Click(Index As Integer)
   
   On Error GoTo FontSelectCancel

   CommonDialog.CancelError = True
   CommonDialog.Color = lblFont(Index).ForeColor
   CommonDialog.FontBold = lblFont(Index).FontBold
   CommonDialog.FontItalic = lblFont(Index).FontItalic
   CommonDialog.FontStrikethru = lblFont(Index).FontStrikethru
   CommonDialog.FontUnderline = lblFont(Index).FontUnderline
   CommonDialog.FontName = lblFont(Index).FontName
   CommonDialog.FONTSIZE = lblFont(Index).FONTSIZE
   CommonDialog.Flags = cdlCFEffects Or cdlCFForceFontExist Or cdlCFPrinterFonts Or cdlCFScalableOnly ' Or cdlCFBoth

   CommonDialog.ShowFont

   lblFont(Index).ForeColor = CommonDialog.Color
   lblFont(Index).FontBold = CommonDialog.FontBold
   lblFont(Index).FontItalic = CommonDialog.FontItalic
   lblFont(Index).FontStrikethru = CommonDialog.FontStrikethru
   lblFont(Index).FontUnderline = CommonDialog.FontUnderline
   lblFont(Index).FontName = CommonDialog.FontName
   lblFont(Index).FONTSIZE = CommonDialog.FONTSIZE

   EnabledStorage

   ' Paint the sample page
   PaintSamplePage

FontSelectCancel:
   CommonDialog.CancelError = False
End Sub

Private Sub EnabledStorage()
   If bBinding Then Exit Sub
   SetEnabled cmdSave, True
   SetEnabled cmdRestore, FileExist(sIniFile)
End Sub

' Factor change value: 765
Private Sub SetSampleOrientation()
   Dim nFactor As Integer, nScale As Integer

   If Printer.Orientation = vbPRORLandscape Then   ' Landscape
      nFactor = 765
      nScale = 270
   Else
      nFactor = 0
      nScale = 0
   End If

   picPage.Width = 2100 + nFactor
   picPage.Height = 2865 - nFactor

   'object.Move Left, Top, Width, Heigh
   lblMM.Move 2565 + nFactor, 3390 - nFactor

   sldLeft.Move 60, 3120 - nFactor, 915 + nScale, 225
   lblLeft(0).Move 150, 3390 - nFactor

   sldRight.Move 1440 + (nFactor - nScale), 3120 - nFactor, 915 + nScale, 225
   lblRight(0).Move 1965 + nFactor, 3390 - nFactor

   sldTop.Move 2265 + nFactor, 165, 225, 1185 - nScale
   lblTop(0).Move 2565 + nFactor, 225

   sldBottom.Move 2265 + nFactor, 2040 - (nFactor - nScale), 225, 1185 - nScale
   lblBottom(0).Move 2565 + nFactor, 2970 - nFactor

End Sub

' Refresh sample now
Private Sub picPage_DblClick()
   If Page.Show Then Exit Sub
   TmrPaint.Enabled = False
   PrintSamplePage
End Sub

' Don't refresh sample just yet - wait for 2 seconds (because it takes a little time to update)
Private Sub PaintSamplePage()
   If bBinding Then Exit Sub
   InvalidateSamplePage
   TmrPaint.Enabled = False
   TmrPaint.Enabled = True
End Sub

' Timer event fired, the 2 seconds must be over - refresh sample page
Private Sub TmrPaint_Timer()
   If Page.Show Then Exit Sub
   TmrPaint.Enabled = False
   PrintSamplePage
End Sub

' --- Status messages (mini help) area -----------------------------------------------------

Private Sub cboExtention_GotFocus()
   
End Sub

Private Sub SetStatusText(sText As String)
   If StatusBar.SimpleText <> sText Then StatusBar.SimpleText = sText
End Sub

Private Sub lstForms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
End Sub

Private Sub cboZoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
End Sub

' ----------------------------------------------------------------------------------------------

Private Sub ClearOutline()
   MdCount = 0
   PrCount = 0
   MdSelected = 0
   PrSelected = 0
   sLoadedFile = ""
   Outline.Clear
   Erase Mdl
   ReDim ItemRef(0)
End Sub

Private Sub SetOutline(sFIle As String, nType As Integer)
   Dim nIndex As Integer, i As Integer, nListIndex As Integer
   Dim sName As String

   nIndex = AnalyseFile(sFIle, nType)

   If nIndex > -1 Then

      sName = Mdl(nIndex).File
      If Mdl(nIndex).ProcCount > 0 Then
         sName = sName & "     [" & Mdl(nIndex).ProcCount & "]"
      End If

      ' List File
      Outline.AddItem sName
      nListIndex = Outline.ListCount - 1
      Outline.Indent(nListIndex) = 1
      Outline.PictureType(nListIndex) = outClosed
      Mdl(nIndex).ListIndex = nListIndex
      MakeReference nListIndex, nIndex, -1

      ' Controls
      If Mdl(nIndex).CtrlCount > 0 Then
         Outline.AddItem "(Controles)" & "     [" & Mdl(nIndex).CtrlCount & "]"
         nListIndex = Outline.ListCount - 1
         Outline.Indent(nListIndex) = 2
         Outline.PictureType(nListIndex) = outClosed
         Mdl(nIndex).CtrlLIndex = nListIndex
         MakeReference nListIndex, nIndex, 0
      End If

      ' Declaration and procedures
      If Mdl(nIndex).ProcCount > 0 Then
         For i = 1 To Mdl(nIndex).ProcCount
            Outline.AddItem Mdl(nIndex).Proc(i).Name
            nListIndex = Outline.ListCount - 1
            Outline.Indent(nListIndex) = 2
            Outline.PictureType(nListIndex) = outClosed
            Mdl(nIndex).Proc(i).ListIndex = nListIndex
            MakeReference nListIndex, nIndex, i
         Next
      End If
   End If

   ShowCounts
   Refresh

End Sub

Private Sub MakeReference(nListIndex As Integer, nFileIndex As Integer, nProcIndex As Integer)
   ReDim Preserve ItemRef(nListIndex)
   ItemRef(nListIndex).FilePoint = nFileIndex
   ItemRef(nListIndex).ProcPoint = nProcIndex      ' -1 = File, 0 = Controls, 1 > = Procedures
End Sub

Private Sub GetProjectDetails()
   On Error Resume Next
   If Outline.ListCount > 0 Then ClearOutline

   If Not FileExist(txtProject) Then
      ButtonsState
      MsgBox "Archivo no encontrado. Seleccione otro archivo", vbInformation
      Exit Sub
   End If

   If txtProject = sLoadedFile Then
      ButtonsState
      Exit Sub
   End If

   Me.MousePointer = vbHourglass
   SetStatusText "Analizando " & ExtractFileName(txtProject) & " ..."
   StatusBar.Refresh

   MakeSound WAVE_ANALYSE

   ' Get the extention to obtain type
   Select Case UCase$(ExtractFileExt(txtProject))
   Case "VBP"
      ' Open the VBP file and extract the files...
      Dim nMark As Integer, nHandle As Integer
      Dim sString As String, sFIle As String, sPath As String
      sPath = ExtractPath(txtProject)

      nHandle = FreeFile
      Open txtProject For Input Access Read Shared As #nHandle
      
      Do While Not EOF(nHandle)  ' Loop until end of file.
         Line Input #nHandle, sString
      
         If UCase$(Left(sString, 4)) = "FORM" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFIle = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFIle, MT_FORM
            End If

         ElseIf UCase$(Left(sString, 6)) = "MODULE" Then
            nMark = InStr(sString, ";")
            If nMark > 0 Then
               sFIle = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFIle, MT_MODULE
            End If

         ElseIf UCase$(Left(sString, 5)) = "CLASS" Then
            nMark = InStr(sString, ";")
            If nMark > 0 Then
               sFIle = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFIle, MT_CLASS
            End If

         ElseIf UCase$(Left(sString, 11)) = "USERCONTROL" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFIle = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFIle, MT_CONTROL
            End If

         ElseIf UCase$(Left(sString, 12)) = "PROPERTYPAGE" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFIle = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFIle, MT_PROPERTY
            End If

         ElseIf UCase$(Left(sString, 12)) = "USERDOCUMENT" Then
            nMark = InStr(sString, "=")
            If nMark > 0 Then
               sFIle = AttachPath(Trim$(Mid$(sString, nMark + 1)), sPath)
               SetOutline sFIle, MT_DOCUMENT
            End If

         End If
      Loop
      
      Close #nHandle

   Case "FRM"
      SetOutline txtProject, MT_FORM

   Case "BAS"
      SetOutline txtProject, MT_MODULE

   Case "CLS"
      SetOutline txtProject, MT_CLASS

   Case "CTL"
      SetOutline txtProject, MT_CONTROL

   Case "PAG"
      SetOutline txtProject, MT_PROPERTY

   Case "DOB"
      SetOutline txtProject, MT_DOCUMENT

   End Select

   sLoadedFile = txtProject
'   AddIniString sIniFile, "Options", "LastFile", txtProject
   UpdateRecentFiles txtProject

   ButtonsState

   SetStatusText "Listo"

   MakeSound WAVE_ACCESSED

   Outline.SetFocus

   Me.MousePointer = vbDefault

End Sub

Private Sub Outline_Click()
   ButtonsState
End Sub

Private Sub CountSelected()
   Dim i As Integer, j As Integer, n As Integer

   MdSelected = 0
   PrSelected = 0

   On Error GoTo SelCountError

   For i = 1 To MdCount
      If Mdl(i).Selected = vbUnchecked Then
         Mdl(i).SelCount = 0

      Else  ' File checked or semi-checked
         n = 0

         If Mdl(i).CtrlSelect = vbChecked Then n = n + 1
         
         ' If file is somehow selected, it can have procedures selected too
         For j = 1 To Mdl(i).ProcCount
            If Mdl(i).Proc(j).Selected = vbChecked Then
               PrSelected = PrSelected + 1
               n = n + 1
            End If
         Next

         If n = 0 Then
            Mdl(i).SelCount = 0
            ' Something must be wrong. No procedures/control selected, but file is?
            Mdl(i).Selected = vbUnchecked
         Else
            Mdl(i).SelCount = n
            MdSelected = MdSelected + 1
         End If
      End If
   Next

SelCountError:
End Sub

Private Sub ShowCounts()
   lblFiles = IIf(MdCount = 0, "Ninguno", MdCount)
   lblProcedures = IIf(PrCount = 0, "Ninguno", PrCount)

   lblSelFiles = IIf(MdSelected = 0, "Ninguno", MdSelected)
   lblSelProcs = IIf(PrSelected = 0, "Ninguno", PrSelected)
End Sub

Private Sub cmdClear_Click()
   Dim i As Integer, j As Integer, nMax As Integer
   On Error Resume Next
   Me.MousePointer = vbHourglass

   If Outline.ListIndex < 0 Or Outline.Indent(Outline.ListIndex) = 1 Then
      ' File... Clear all entries

      For i = 0 To MdCount
         Mdl(i).Selected = vbUnchecked
         Mdl(i).CtrlSelect = vbUnchecked

         ' If file is somehow selected, it must have procedures selected too
         For j = 1 To Mdl(i).ProcCount
            Mdl(i).Proc(j).Selected = vbUnchecked
         Next
      Next

      ' Remove all "checked" boxes
      For i = 0 To (Outline.ListCount - 1)
         Outline.PictureType(i) = outClosed
      Next

   Else
      ' Procedure... Clear relatives only
      i = ItemRef(Outline.ListIndex).FilePoint     ' Get the file (parent) array pointer
      SetChildrenTick i, vbUnchecked

      Mdl(i).Selected = vbUnchecked
      Outline.PictureType(Mdl(i).ListIndex) = outClosed
   End If

   CountSelected
   ButtonsState
   ShowCounts
   Me.MousePointer = vbDefault
End Sub

Private Sub cmdSelectAll_Click()
   Dim i As Integer, j As Integer
   On Error Resume Next
   Me.MousePointer = vbHourglass

   If Outline.ListIndex < 0 Or Outline.Indent(Outline.ListIndex) = 1 Then
      ' File... Clear all entries

      For i = 0 To MdCount
         Mdl(i).Selected = vbChecked
         Mdl(i).CtrlSelect = vbChecked

         ' If file is somehow selected, it must have procedures selected too
         For j = 1 To Mdl(i).ProcCount
            Mdl(i).Proc(j).Selected = vbChecked
         Next
      Next

      ' Set all boxes to "checked"
      For i = 0 To (Outline.ListCount - 1)
         Outline.PictureType(i) = outOpen
      Next

   Else
      ' Procedure... Set relatives only
      i = ItemRef(Outline.ListIndex).FilePoint     ' Get the file (parent) array pointer
      SetChildrenTick i, vbChecked

      Mdl(i).Selected = vbChecked
      Outline.PictureType(Mdl(i).ListIndex) = outOpen
   End If
   
   CountSelected
   ButtonsState
   ShowCounts
   Me.MousePointer = vbDefault
End Sub

Private Sub Outline_PictureClick(ListIndex As Integer)
   Outline.MousePointer = vbHourglass

   Select Case ItemRef(ListIndex).ProcPoint
   Case Is < 0
      ' It's a file (parent)...
      If Mdl(ItemRef(ListIndex).FilePoint).Selected = vbUnchecked Then
         Mdl(ItemRef(ListIndex).FilePoint).Selected = vbChecked
         SetChildrenTick ItemRef(ListIndex).FilePoint, vbChecked     ' Select all children too
         Outline.PictureType(ListIndex) = outOpen
      Else
         Mdl(ItemRef(ListIndex).FilePoint).Selected = vbUnchecked
         SetChildrenTick ItemRef(ListIndex).FilePoint, vbUnchecked   ' Unselect all children too
         Outline.PictureType(ListIndex) = outClosed
      End If

   Case Is = 0
      ' It's the controls...
      If Mdl(ItemRef(ListIndex).FilePoint).CtrlSelect = vbUnchecked Then
         Mdl(ItemRef(ListIndex).FilePoint).CtrlSelect = vbChecked
         Outline.PictureType(ListIndex) = outOpen
      Else
         Mdl(ItemRef(ListIndex).FilePoint).CtrlSelect = vbUnchecked
         Outline.PictureType(ListIndex) = outClosed
      End If
      SetParentTick ItemRef(ListIndex).FilePoint

   Case Is > 0
      ' It's a procedure or declaration...
      If Mdl(ItemRef(ListIndex).FilePoint).Proc(ItemRef(ListIndex).ProcPoint).Selected = vbUnchecked Then
         Mdl(ItemRef(ListIndex).FilePoint).Proc(ItemRef(ListIndex).ProcPoint).Selected = vbChecked
         Outline.PictureType(ListIndex) = outOpen
      Else
         Mdl(ItemRef(ListIndex).FilePoint).Proc(ItemRef(ListIndex).ProcPoint).Selected = vbUnchecked
         Outline.PictureType(ListIndex) = outClosed
      End If
      SetParentTick ItemRef(ListIndex).FilePoint

   End Select

   CountSelected
   ButtonsState
   ShowCounts
   Outline.MousePointer = vbDefault
End Sub

Private Sub SetParentTick(nFileIndex As Integer)
   Dim nImage As Integer, i As Integer

   ' Get the first child's status
   If Mdl(nFileIndex).CtrlCount > 0 Then
      ' Controls is considered a child.
      nImage = IIf(Mdl(nFileIndex).CtrlSelect = vbChecked, outOpen, outClosed)
   ElseIf Mdl(nFileIndex).ProcCount > 0 Then
      nImage = IIf(Mdl(nFileIndex).Proc(1).Selected = vbChecked, outOpen, outClosed)
   Else
      nImage = IIf(Mdl(nFileIndex).Selected = vbChecked, outOpen, outClosed)
   End If

   If Mdl(nFileIndex).ProcCount > 0 Then
      For i = 1 To Mdl(nFileIndex).ProcCount
         If Mdl(nFileIndex).Proc(i).Selected = vbChecked And nImage = outClosed Then
            nImage = outLeaf     ' Grey it...
            Exit For
         ElseIf Mdl(nFileIndex).Proc(i).Selected = vbUnchecked And nImage = outOpen Then
            nImage = outLeaf     ' Grey it...
            Exit For
         End If
      Next
   End If

   Select Case nImage
   Case outOpen, outLeaf
      Mdl(nFileIndex).Selected = True
   Case outClosed
      Mdl(nFileIndex).Selected = False
   End Select

   Outline.PictureType(Mdl(nFileIndex).ListIndex) = nImage
End Sub

Private Sub SetChildrenTick(nFileIndex As Integer, nSelect As Integer)
   Dim nImage As Integer, i As Integer
   nImage = IIf(nSelect = vbChecked, outOpen, outClosed)

   If Mdl(nFileIndex).CtrlCount > 0 Then
      ' Controls item is considered a child.
      Mdl(nFileIndex).CtrlSelect = nSelect
      Outline.PictureType(Mdl(nFileIndex).CtrlLIndex) = nImage
   End If

   If Mdl(nFileIndex).ProcCount > 0 Then
      For i = 1 To Mdl(nFileIndex).ProcCount
         Mdl(nFileIndex).Proc(i).Selected = nSelect
         Outline.PictureType(Mdl(nFileIndex).Proc(i).ListIndex) = nImage
      Next
   End If
End Sub

Private Sub ButtonsState()
   On Error Resume Next

   Dim nIndex As Integer
   If Outline.ListCount = 0 Then
      nIndex = -1
   Else
      nIndex = Outline.ListIndex
   End If

   ShowCounts

   If Outline.ListCount < 1 Or nIndex < 0 Then
      lblName = "(No hay item seleccionado)"
      lblType = ""
   Else
      Select Case ItemRef(nIndex).ProcPoint
      Case Is < 0
         ' File...
         Select Case Mdl(ItemRef(nIndex).FilePoint).Type
         Case MT_FORM
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Form"
            Else
               lblType = "Form - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_MODULE
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Modulo"
            Else
               lblType = "Modulo - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_CLASS
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Clase"
            Else
               lblType = "Clase - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_CONTROL
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Control de Usuario"
            Else
               lblType = "Control de Usuario - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_PROPERTY
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Pagina de Propiedades"
            Else
               lblType = "Pagina de Propiedades - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case MT_DOCUMENT
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            If Len(Trim(Mdl(ItemRef(nIndex).FilePoint).Name)) = 0 Then
               lblType = "Documento de Usuario"
            Else
               lblType = "Documento de Usuario - " & Mdl(ItemRef(nIndex).FilePoint).Name
            End If
         Case Else
            ' Dunno
            lblName = ""
            lblType = ""
         End Select

      Case Is = 0
         ' Controls
         If Mdl(ItemRef(nIndex).FilePoint).CtrlCount > 0 Then
            lblName = LongDirFix(Mdl(ItemRef(nIndex).FilePoint).PathFile, 30)
            lblType = "Controles"
         Else
            lblName = ""
            lblType = ""
         End If
      
      Case Is > 0
         ' Procedure
         Select Case Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Type
         Case PT_DECLARE
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Declaraciones"
         Case PT_PROPERTY
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Propiedades"
         Case PT_SUB
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Sub"
         Case PT_FUNCTION
            lblName = Mdl(ItemRef(nIndex).FilePoint).Proc(ItemRef(nIndex).ProcPoint).Name
            lblType = "Function"
         Case Else
            ' Dunno
            lblName = ""
            lblType = ""
         End Select

      End Select
   End If

   If Outline.ListCount > 0 Then

      SetEnabled cmdView, (nIndex > -1)
      tbrMain.Buttons("cmdPreview").Enabled = cmdView.Enabled
      
      If nIndex < 0 Then
         SetEnabled cmdSelectAll, (MdCount <> MdSelected)
         tbrMain.Buttons("cmdSelTodo").Enabled = cmdSelectAll.Enabled
         SetEnabled cmdClear, (MdSelected > 0)
         tbrMain.Buttons("cmdLimTodo").Enabled = cmdClear.Enabled
      ElseIf Outline.Indent(nIndex) = 1 Then
         SetEnabled cmdSelectAll, (MdCount <> MdSelected)
         tbrMain.Buttons("cmdSelTodo").Enabled = cmdSelectAll.Enabled
         SetEnabled cmdClear, (MdSelected > 0)
         tbrMain.Buttons("cmdLimTodo").Enabled = cmdClear.Enabled
      Else
         SetEnabled cmdSelectAll, (Mdl(ItemRef(nIndex).FilePoint).ChildCount <> Mdl(ItemRef(nIndex).FilePoint).SelCount)
         tbrMain.Buttons("cmdSelTodo").Enabled = cmdSelectAll.Enabled
         SetEnabled cmdClear, (Mdl(ItemRef(nIndex).FilePoint).SelCount > 0)
         tbrMain.Buttons("cmdLimTodo").Enabled = cmdClear.Enabled
      End If

      If MdSelected > 0 Then
         SetEnabled cmdPrint, True
         tbrMain.Buttons("cmdImprimir").Enabled = True
      Else
         SetEnabled cmdPrint, (chkProject = vbChecked And UCase$(ExtractFileExt(txtProject)) = "VBP")
         tbrMain.Buttons("cmdImprimir").Enabled = cmdPrint.Enabled
      End If

   Else
      SetEnabled cmdView, False
      tbrMain.Buttons("cmdPreview").Enabled = cmdView.Enabled
      SetEnabled cmdClear, False
      SetEnabled cmdSelectAll, False
      tbrMain.Buttons("cmdSelTodos").Enabled = False
      SetEnabled cmdPrint, (chkProject = vbChecked And UCase$(ExtractFileExt(txtProject)) = "VBP")
      tbrMain.Buttons("cmdImprimir").Enabled = cmdPrint.Enabled
   End If

   SetEnabled cmdPrintSetup, (Not optOutput(2))
   SetEnabled cmdHelp, FileExist(sHelpFile)

End Sub

' Some comments in the footer
' I don't know why, but it's here - What to do with it?
' It actually belongs to the procedure above, but how to connect to it without producing a line
' or page break. - I do need something to test this program with, why not this.
