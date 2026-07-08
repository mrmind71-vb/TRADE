VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form StoreData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáãÎÇÒä"
   ClientHeight    =   6315
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   7290
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   7290
   Begin VB.CheckBox XISSTOREBRANCH 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáÈíÇäÇÊ ONLINE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   1395
      Width           =   1905
   End
   Begin VB.CheckBox XISNOGRDISC 
      Alignment       =   1  'Right Justify
      Caption         =   "ÈÏæä ãÌãæÚÇÊ ÎÕã"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3780
      Width           =   2265
   End
   Begin VB.TextBox XDATE2_BR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   3150
      Width           =   1500
   End
   Begin VB.TextBox XDATE1_BR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   2745
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   4590
      Width           =   6855
      Begin VB.CommandButton Command1 
         Caption         =   "ÊÍÏíË ãÈíÚÇÊ ÇáãäÇØÞ Úáì ÇáÊæÒíÚ ÇáÍÇáì ááãÏíÑíä"
         Height          =   735
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   2445
      End
      Begin VB.TextBox XDATE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3105
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   630
         Width           =   1860
      End
      Begin VB.TextBox XDATE1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3105
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Åáì ÊÇÑíÎ : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "ãä ÊÇÑíÎ : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   5085
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.CheckBox xonline 
      Alignment       =   1  'Right Justify
      Caption         =   "Çæä áÇíä"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4185
      Width           =   1455
   End
   Begin VB.CheckBox XISONEST2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÝÑÚ Êæßíá "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4230
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3735
      Width           =   1230
   End
   Begin VB.CheckBox XISONEST 
      Alignment       =   1  'Right Justify
      Caption         =   "ÝÑÚ ÇãÇäÇÊ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2452
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4185
      Width           =   1455
   End
   Begin VB.CheckBox xisstop 
      Alignment       =   1  'Right Justify
      Caption         =   "ÝÑÚ  ãÊæÞÝ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4185
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   7290
      TabIndex        =   5
      Top             =   5850
      Width           =   7290
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   4425
         Picture         =   "StoreData.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "ÇáÊÇáí"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   3510
         Picture         =   "StoreData.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "ÇáÓÇÈÞ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   2355
         Picture         =   "StoreData.frx":4B93
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "ÃÎíÑ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   1485
         Picture         =   "StoreData.frx":726D
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ãæá"
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4500
      MaxLength       =   3
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   1005
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   855
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   985
      Width           =   4650
   End
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7290
      TabIndex        =   2
      Top             =   0
      Width           =   7290
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData.frx":994C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "ÎÑæÌ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData.frx":9A96
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "ÍÝÙ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData.frx":9ED8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "ÍÐÝ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData.frx":C772
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "ÊÑÇÌÚ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   4740
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData.frx":ECEB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "ÇÖÇÝÉ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   5880
         Picture         =   "StoreData.frx":11297
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "ÇÓÊÚáÇã"
         Top             =   45
         Width           =   1140
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo XBRANCH 
      Height          =   390
      Left            =   2610
      TabIndex        =   17
      Top             =   1385
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo xcust 
      Height          =   390
      Left            =   855
      TabIndex        =   20
      Top             =   1845
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo XBRANCH_GR 
      Height          =   390
      Left            =   855
      TabIndex        =   22
      Top             =   2295
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   810
      Top             =   585
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo xacc 
      Height          =   390
      Left            =   3060
      TabIndex        =   25
      Top             =   2745
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   45
      Top             =   585
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo xacc2 
      Height          =   390
      Left            =   3060
      TabIndex        =   34
      Top             =   3195
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ãÕÇÑíÝ ÞÈá ÊÇÑíÎ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1620
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   3240
      Width           =   1410
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ãÕÇÑíÝ ÈÛÏ ÊÇÑíÎ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   2790
      Width           =   1410
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ÍÓÇÈ Êæßíá ÓÇÈÞ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3285
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ÍÓÇÈ ÇáÊæßíá "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2835
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáãäØÞÉ : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2385
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÚãíá :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1935
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÝÑÚ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáãÎÒä "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ßæÏ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   660
      Width           =   315
   End
End
Attribute VB_Name = "StoreData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer, bedit As Boolean
Dim con As New ADODB.Connection
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim CardTable As New ADODB.Recordset
Const LoadMode = 1, DefineMode = 2

Private Sub Command1_Click()
On Error GoTo myerror
If MsgBox(" ÊÍÏíË ÇáãäÇØÞ Úáì ÇáÝÑæÚ  ÎáÇá ÇáÞäÑÉ ÍÓÈ ÇáÊæÒíÚ ÇáÍÇáì", vbYesNo + vbDefaultButton2) = vbYes Then
    con.CommandTimeout = 5000
    con.Execute "    UPDATE FILE6_20H SET   GR_MAN = FILE0_40.BRANCH_GR  FROM FILE6_20H INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE DATE >= " & DateSq(XDATE1) & " AND DATE <= " & DateSq(XDATE2.text)
    con.Execute "    UPDATE FR6_20H         SET GR_MAN = BRANCH_FR.BRANCH_GR  FROM FR6_20H INNER JOIN BRANCH_FR ON BRANCH_FR.CODE = FR6_20H.STORE WHERE DATE >= " & DateSq(XDATE1) & " AND DATE <= " & DateSq(XDATE2.text)
    MsgBox " Êã ÊÚÏíá ãÈíÚÇÊ ÇáãäÇØÞ ááÝÊÑÉ  "
End If
Exit Sub
myerror:
Err.Clear
End Sub

Private Sub Form_Load()
openCon con
bedit = True
CmdDel.Visible = False

Frame1.Visible = bSupermode

data1.ConnectionString = strCon
data1.RecordSource = "Select * From BRANCH"
Set XBRANCH.RowSource = data1
XBRANCH.ListField = "Desca"
XBRANCH.BoundColumn = "Code"

data2.ConnectionString = strCon
data2.RecordSource = "Select * From file3_10 order by desca "
Set xcust.RowSource = data2
xcust.ListField = "Desca"
xcust.BoundColumn = "Code"

DATA3.ConnectionString = strCon
DATA3.RecordSource = "Select * From BRANCH_GR order by CODE "
Set XBRANCH_GR.RowSource = DATA3
XBRANCH_GR.ListField = "Desca"
XBRANCH_GR.BoundColumn = "Code"

DATA4.ConnectionString = strCon
DATA4.RecordSource = "Select * From ACC0_10 "
Set xacc.RowSource = DATA4
xacc.ListField = "Desca"
xacc.BoundColumn = "Code"

Set xacc2.RowSource = DATA4
xacc2.ListField = "Desca"
xacc2.BoundColumn = "Code"

CardTable.Open "SELECT * FROM FILE0_40  ORDER BY CODE", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdAdd_Click()
myDefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("ÇáÛÇÁ ÇáÓÌá ÇáÍÇáì : åá ÇäÊ ãæÇÝÞ ¿", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From FILE0_40 Where code = " & MyParn(xCode.text)
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        myDefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub

Inform "Êã ÍÝÙ ÇáÈíÇäÇÊ ÈäÌÇÍ"
CardTable.Requery
If xCode.Enabled Then
    CmdAdd_Click
Else
    CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xCode.Enabled Then
        myDefine
    Else
        CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub CmdFirst_Click()
    CardTable.MoveFirst
    myload
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub CmdLast_Click()
    CardTable.MoveLast
    myload
End Sub
Private Sub CmdNext_Click()
    CardTable.MoveNext
    If CardTable.EOF Then
        CardTable.MovePrevious
    Else
        myload
    End If
End Sub
Private Sub CmdPrevious_Click()
    CardTable.MovePrevious
    If CardTable.BOF Then
        CardTable.MoveNext
    Else
        myload
    End If
End Sub
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
CmdSave.Enabled = bedit
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE0_40 "
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "ÇáÈíÇä"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "ÇáßæÏ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ÇáÈíÇä"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "ÅÓÊÚáÇã "
Search3.Show 1
End Sub
Sub myDefine()

    xCode.text = RetZero(Val(Newflag("FILE0_40", "CODE", con)), 3)
    xDescA.text = ""
    XBRANCH.BoundText = ""
    xacc.BoundText = ""
    xacc2.BoundText = ""
    XBRANCH_GR.BoundText = ""
    xcust.BoundText = ""
    XISONEST.Value = 0
    XISONEST2.Value = 0
    xisstop.Value = 0
    XISSTOREBRANCH.Value = 0
    XDATE1_BR.text = ""
    XDATE2_BR.text = ""
    XISNOGRDISC.Value = 0
    xonline.Value = 0
    Handlecontrols DefineMode

End Sub
Sub myload()
xCode.text = CardTable!CODE & ""
xDescA.text = CardTable!DESCA & ""
xisstop.Value = IIf(CardTable!ISSTOP, 1, 0)
xonline.Value = IIf(CardTable!online, 1, 0)
XISONEST.Value = IIf(CardTable!ISONEST, 1, 0)
XISONEST2.Value = IIf(CardTable!ISONEST2, 1, 0)
XISNOGRDISC.Value = IIf(CardTable!ISNOGRDISC, 1, 0)
XISSTOREBRANCH.Value = IIf(CardTable!ISSTOREBRANCH, 1, 0)
XDATE1_BR.text = CardTable!date1_br & ""
XDATE2_BR.text = CardTable!date2_br & ""

XBRANCH.BoundText = CardTable!branch & ""
xacc.BoundText = CardTable!ACC & ""
xacc2.BoundText = CardTable!ACC2 & ""
xcust.BoundText = CardTable!CUST & ""
XBRANCH_GR.BoundText = CardTable!branch_GR & ""
XBRANCH_GR.BoundText = CardTable!branch_GR & ""
xacc.BoundText = CardTable!ACC & ""
xRecordNumber = "ÓÌá " & CardTable.AbsolutePosition + 1 & " ãä " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(14, 1)

aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "DESCA"
aInsert(1, 1) = addstring(xDescA.text)

aInsert(2, 0) = "BRANCH"
aInsert(2, 1) = addstring(XBRANCH.BoundText)

aInsert(3, 0) = "ISONEST"
aInsert(3, 1) = IIf(XISONEST.Value, 1, 0)

aInsert(4, 0) = "ISSTOP"
aInsert(4, 1) = IIf(xisstop.Value, 1, 0)

aInsert(5, 0) = "CUST"
aInsert(5, 1) = addstring(xcust.BoundText)

aInsert(6, 0) = "BRANCH_GR"
aInsert(6, 1) = addvalue(XBRANCH_GR.BoundText)

aInsert(7, 0) = "ACC"
aInsert(7, 1) = addstring(xacc.BoundText)

aInsert(8, 0) = "ISONEST2"
aInsert(8, 1) = IIf(XISONEST2.Value, 1, 0)

aInsert(9, 0) = "online"
aInsert(9, 1) = IIf(xonline.Value, 1, 0)

aInsert(10, 0) = "date1_br"
aInsert(10, 1) = addDate(XDATE1_BR.text)

aInsert(11, 0) = "date2_br"
aInsert(11, 1) = addDate(XDATE2_BR.text)

aInsert(12, 0) = "ACC2"
aInsert(12, 1) = addstring(xacc2.BoundText)

aInsert(13, 0) = "ISNOGRDISC"
aInsert(13, 1) = IIf(XISNOGRDISC.Value, 1, 0)

aInsert(14, 0) = "ISSTOREBRANCH"
aInsert(14, 1) = IIf(XISSTOREBRANCH.Value, 1, 0)

On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    con.Execute CreateInsert(aInsert, "FILE0_40")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_40", " where CODE = " & addstring(xCode.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
   CardTable.Find "CODE = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
   myload
   Search3.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Unload Search3
Set Search3 = Nothing
Err.Clear
closeCon con
End Sub
Private Sub xCode_LostFocus()
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 2)
CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "ÇáßæÏ áÇ íãßä Çä íßæä ÎÇáíÇ"
    Exit Function
End If

If xDescA.text = "" Then
    MsgBox "ÇáÅÓã áÇ íãßä Çä íßæä ÎÇáíÇ"
    Exit Function
End If

MYVALID = True
End Function
Private Sub xCode_GotFocus()
    xCode.SelStart = 0
    xCode.SelLength = Len(xCode.text)
End Sub
Private Sub xDescA_GotFocus()
    xDescA.SelStart = 0
    xDescA.SelLength = Len(xDescA.text)
End Sub
