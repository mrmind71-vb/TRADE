VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ModelMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… „ÊœÌ·"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   0
      Width           =   2985
      Begin VB.Label XBALCOST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   585
         Width           =   1725
      End
      Begin VB.Label XBAL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁÌ„… : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   3060
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1080
      Width           =   3660
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2430
         Picture         =   "ModelMove.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   1230
         Picture         =   "ModelMove.frx":3059
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "ModelMove.frx":5483
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   3060
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   3660
      Begin VB.TextBox xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   90
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   540
      Top             =   -135
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
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   8385
      Begin VB.TextBox xItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4815
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1575
         Width           =   2310
      End
      Begin VB.TextBox xModel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4815
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2310
      End
      Begin MSDataListLib.DataCombo xscal 
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Top             =   885
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xColor 
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         Top             =   1230
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xcode 
         Height          =   315
         Left            =   3105
         TabIndex        =   26
         Top             =   540
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   540
         Width           =   495
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   4695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·»«—ﬂÊœ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1620
         Width           =   1035
      End
      Begin VB.Label xSupp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1230
         Width           =   2985
      End
      Begin VB.Label xFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1230
         Width           =   2580
      End
      Begin VB.Label xMosm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   885
         Width           =   2985
      End
      Begin VB.Label xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   885
         Width           =   2580
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "·Ê‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1260
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„ﬁ«” :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1035
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   405
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1800
      Top             =   45
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   90
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   24
      Top             =   9390
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "02:02 „"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6885
      Left            =   45
      TabIndex        =   33
      Top             =   2115
      Width           =   15090
      _cx             =   26617
      _cy             =   12144
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16776960
      ForeColorSel    =   64
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   90
      Top             =   1395
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
End
Attribute VB_Name = "ModelMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearchItem As New Search3
Dim con As New ADODB.Connection
Dim sitem As Double
Sub Fillgrd()
Dim loctable As New ADODB.Recordset
cString = "select SUM(FILE1_11.COST * ( [IN] + [OUT] ) ) AS TCOST , file1_11.date ,FILE1_12.DESCA,file1_11.doc_id , file1_11.store , file1_11.[type] , sum(file1_11.[in]) as tin , sum(file1_11.out) as tout  ,file1_12.desca AS TYPEDESCA  ,file4_10.desca as supDesca,file0_40.desca as storeDesca , SUM(FILE1_11.TOTAL ) as TOTAL , FILE1_12.[ORDER] " & _
          " FROM  (((FILE1_11 INNER JOIN  FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM)  LEFT JOIN FILE1_12 ON FILE1_11.TYPE = FILE1_12.CODE) INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE) LEFT JOIN FILE4_10 ON FILE1_11.codesup = FILE4_10.CODE WHERE FILE1_11.ITEM  IS NOT NULL "

If IsNumeric(xItem.text) Then
    cString = cString & " AND file1_10.ITEM = " & xItem.text
Else
    If xModel.text <> "" Then
        cString = cString & " AND file1_10.MODELNO = " & MyParn(xModel.text)
    End If
    
    If IsDate(xdate.text) Then
        cString = cString & " AND file1_11.date >= " & DateSq(xdate.text)
    End If
    
    
    If Trim(xcode.BoundText) <> "" Then
        cString = cString & " AND file1_10.CODE = " & MyParn(xcode.BoundText)
    End If
    
    If Trim(xscal.BoundText) <> "" Then
        cString = cString & " AND file1_10.SCAL = " & MyParn(xscal.text)
    End If
    
    If Trim(xColor.BoundText) <> "" Then
        cString = cString & " AND file1_10.COLOR = " & MyParn(xColor.text)
    End If
End If

If Trim(xStore.BoundText) <> "" Then
    cString = cString & " AND file1_11.store = " & MyParn(xStore.BoundText)
End If

cString = cString & " group by file1_11.date , file1_11.doc_id , file1_11.store , file1_11.[type]  ,file1_12.desca ,file4_10.desca ,file0_40.desca  , FILE1_12.[ORDER] "
cString = cString & cWhere & " ORDER BY FILE1_11.DATE, FILE1_12.[ORDER] "

With grid1
    .Rows = 1
    If IsDate(xdate.text) Then
       cString2 = "Select sum([IN] - OUT ) as Balance from file1_11 where file1_11.item = " & MyParn(xItem.text) & _
                  " and file1_11.date < " & DateSq(xdate.text) & cWhere
       nPrevious = Val(GetDesca(cString2, con))
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate.text
            .TextMatrix(.Rows - 1, 3) = nPrevious
       End If
    End If
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Val(loctable!tIn & "") - Val(loctable!tout & "")
        .TextMatrix(.Rows - 1, 0) = loctable!Desca & ""
        .TextMatrix(.Rows - 1, 1) = Format(Val(loctable!tout & ""), "#0")
        .TextMatrix(.Rows - 1, 2) = Format(Val(loctable!tIn & ""), "#0")
        .TextMatrix(.Rows - 1, 3) = Format(nPrevious, "#0")
        .TextMatrix(.Rows - 1, 4) = Format(loctable!Date, "yyyy/mm/dd")
        If (loctable!Type = "6" Or cBranch = "00") And (loctable!tout + loctable!tIn) <> 0 Then
            .TextMatrix(.Rows - 1, 5) = (Val(loctable!TOTAL & "") / (loctable!tout + loctable!tIn))
        End If
        .TextMatrix(.Rows - 1, 6) = loctable!doc_id & ""
        .TextMatrix(.Rows - 1, 7) = loctable!SUPDESCA & " " & loctable!storedesca
        .TextMatrix(.Rows - 1, 8) = loctable!Type & ""
        If loctable!Type = "2" Or loctable!Type = "7" Then
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbGreen
        End If
        If (loctable!tout + loctable!tIn) <> 0 Then
            .TextMatrix(.Rows - 1, 9) = Round(loctable!tcost / (loctable!tout + loctable!tIn), 2) & ""
        End If
        If loctable!Type = "6" And loctable!tout <> 0 Then
            If loctable!tout <> 0 Then .TextMatrix(.Rows - 1, 10) = Round(Val(loctable!TOTAL & "") / loctable!tout, 2)
        End If
        loctable.MoveNext
    Loop
    XBAL.Caption = Val(.TextMatrix(.Rows - 1, 3))
    XBALCOST.Caption = Round(Val(.TextMatrix(.Rows - 1, 9)) * Val(XBAL.Caption), 2)
End With
End Sub
Sub myProc()
    ActiveControl.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    oSearchItem.Hide
End Sub
Function MYVALID() As Boolean
If Trim(xModel.text) = "" And Trim(xItem.text) = "" Then
    MsgBox "€Ì— „”Ã· „ÊœÌ· «Ê »«—ﬂÊœ"
    Exit Function
End If
If (Not IsDate(xdate.text)) And Trim(xdate.text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If Not bOpt5 And xStore.BoundText = "" Then
    MsgBox " ÕœÌœ ›—⁄"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdGo_Click()
If Not MYVALID Then Exit Sub
If Trim(xItem.text) <> "" Then
    Dim loctable As ADODB.Recordset
    Set loctable = ItemFind(xItem.text, con)
    If Not (loctable.EOF And loctable.BOF) Then
        xdesca.Caption = loctable!Desca
        xcode.BoundText = loctable!CODE
        xModel.text = loctable!MODEL
        xModelFact.Caption = loctable!modelfact0
        xscal.text = loctable!scal
        xColor.text = loctable!color
    End If
End If
Fillgrd
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CMD_PRINT_Click()
'    Load rpitem2
'    rpitem2.xItem.Text = xItem.Text
'    rpitem2.xdate1.Text = xdate.Text
'    rpitem2.xStore.BoundText = xStore.BoundText
'    rpitem2.Show 1
    
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String

    If xSupp.Caption = "" Then
        cHead1 = ArbString(" ›’Ì·Ï Õ—ﬂ… „ÊœÌ· —ﬁ„ " & xModelFact.Caption & " „’‰⁄ " & xFact.Caption & " „Ê”„ " & xMosm.Caption)
    Else
        cHead1 = ArbString(" ›’Ì·Ï Õ—ﬂ… „ÊœÌ· —ﬁ„ " & xModelFact.Caption & " „’‰⁄ " & xFact.Caption & " „Ê”„ " & xMosm.Caption & " „ﬂ » " & xSupp.Caption)
    End If
    If xscal.BoundText <> "" Then cHead2 = cHead2 & ArbString(" „ﬁ«” " & xscal.text)
    If xColor.BoundText <> "" Then cHead2 = cHead2 & ArbString(" ·Ê‰ " & xColor.text)
    If xStore.BoundText <> "" Then cHead3 = cHead3 & ArbString(" „Œ“‰  " & xStore.text)
    If IsDate(xdate.text) Then cHead3 = cHead3 & ArbString(" „‰  «—ÌŒ " & xdate.text)
    
    Load PrintGrd
    PrintGrd.doprint grid1, 0.8, 1, cHead1, cHead2, cHead3, False, , 9
    PrintGrd.Show 1

End Sub

Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -2, , , , False, True, 6
    PrintGrd.Show 1

End Sub

Private Sub Form_Load()
openCon con
Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE " & cUserStore, con)
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

If cBranch <> "00" Then
    xStore.BoundText = sStore
    xStore.Enabled = False
End If

Set data4.Recordset = myRecordSet("SELECT * FROM FILE4_10 order by desca ", con)
Set xcode.RowSource = data4
xcode.ListField = "Desca"
xcode.BoundColumn = "Code"

With grid1
.FixedCols = 0

grid1.Cols = 11
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = "’«œ—"
.TextMatrix(0, 2) = "Ê«—œ"
.TextMatrix(0, 3) = "—’Ìœ"
.TextMatrix(0, 4) = " «—ÌŒ"
.TextMatrix(0, 5) = "«·ﬁÌ„…"
.TextMatrix(0, 6) = "„” ‰œ"
.TextMatrix(0, 7) = "≈”„"
.TextMatrix(0, 9) = " ﬂ·›…"
.TextMatrix(0, 10) = "’«›Ï ”⁄—"
.ColHidden(9) = Not bOpt10
.ColHidden(8) = True
grid1.ColWidth(0) = 2500
grid1.ColWidth(1) = 1000
grid1.ColWidth(2) = 1000
grid1.ColWidth(3) = 1000
grid1.ColWidth(4) = 1500
grid1.ColWidth(5) = 1000
grid1.ColWidth(6) = 1700
grid1.ColWidth(7) = 2000
grid1.ColWidth(9) = 1000
grid1.ColWidth(10) = 1000
'If cBranch <> "00" Then .ColHidden(5) = True
End With
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next


Set DATA2.Recordset = myRecordSet("SELECT  C_SCAL , SCAL FROM FILE1_10 WHERE 1 = 2", con)
Set xscal.RowSource = DATA2
xscal.ListField = "SCAL"
xscal.BoundColumn = "C_SCAL"

Set DATA3.Recordset = myRecordSet("SELECT C_COLOR, COLOR FROM FILE1_10 WHERE 1 = 2", con)
xColor.ListField = "COLOR"
xColor.BoundColumn = "C_COLOR"
Set xColor.RowSource = DATA3

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload oSearchItem
Err.Clear
End Sub

Private Sub CalcScal(sModel)
Dim cString As String, sValue As String
cValue = xscal.BoundText
If Trim(sModel) = "" Then
    cString = "SELECT  C_SCAL , SCAL FROM FILE1_10 WHERE 1 = 2"
Else
    cString = "SELECT  C_SCAL , SCAL FROM FILE1_10"
    cString = cString & turn(cString) & "MODELNO = " & MyParn(sModel)
    cString = cString & " GROUP BY C_SCAL , SCAL ORDER BY C_SCAL "
End If
Set DATA2.Recordset = myRecordSet(cString, con)
xscal.BoundText = sValue
If Not xscal.MatchedWithList Then xscal.BoundText = ""
End Sub
Private Sub CalcColor(sModel)
Dim cString As String, sValue As String
cValue = xColor.BoundText
If Trim(sModel) = "" Then
    cString = "SELECT  C_COLOR , COLOR FROM FILE1_10 WHERE 1 = 2"
Else
    cString = "SELECT  C_COLOR , COLOR FROM FILE1_10"
    cString = cString & turn(cString) & "MODELNO = " & MyParn(sModel)
    cString = cString & " GROUP BY C_COLOR , COLOR ORDER BY C_COLOR "
End If
Set DATA3.Recordset = myRecordSet(cString, con)
xColor.BoundText = sValue
If Not xColor.MatchedWithList Then xColor.BoundText = ""
End Sub
Private Sub xItem_Validate(Cancel As Boolean)
If IsNumeric(xItem.text) Then
    If Trim(xModel.text) <> "" Then
        xModel.text = ""
        XModel_LostFocus
    End If
Else
    If Trim(xModel.text) = "" Then xdesca.Caption = ""
End If
End Sub
Private Sub xModel_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then ModelNoLookupAll Me, oSearchItem
End Sub
Private Sub XModel_LostFocus()
If Trim(xModel.text) = "" Then
    xdesca.Caption = ""
    xFact.Caption = ""
    xMosm.Caption = ""
    xModelFact.Caption = ""
    xSupp.Caption = ""
Else
    Dim cString As String, aRet As Variant
    cString = "SELECT FILE1_10H.DESCA,FACT, MOSM , MODELFACT , SUPP , FACT.DESCA FROM FILE1_10H INNER JOIN FACT ON FACT.CODE = FILE1_10H.FACT "
    cString = cString & turn(cString) & "MODELNO = " & MyParn(xModel.text)
    aRet = aGetDesca(cString, con)
    If UBound(aRet) > 0 Then
        xdesca.Caption = aRet(1)
        xFact.Caption = aRet(6) ' GetDesca("SELECT DESCA FROM FACT WHERE CODE = " & MyParn(aRet(2)))
        xMosm.Caption = aRet(3)
        xModelFact.Caption = DelZero(aRet(4))
        xSupp.Caption = aRet(5) & ""
        CalcColor xModel.text
        CalcScal xModel.text
        CalcCode xModel.text
    Else
        CalcColor ""
        CalcScal ""
        CalcCode ""
    End If
End If
End Sub
Private Sub xStore_Click(Area As Integer)
If Not cmdGo.Enabled Then cmdGo.Enabled = True
End Sub
Private Sub xStore_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdGo_Click
End Sub
Private Sub xstore_LostFocus()
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub CalcCode(sModel)
Dim cString As String, sValue As String
cValue = xcode.BoundText
If Trim(sModel) = "" Then
    cString = "SELECT  CODE , DESCA FROM FILE4_10 WHERE 1 = 2"
Else
    cString = "SELECT  FILE4_10.CODE , FILE4_10.DESCA FROM FILE1_10 INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE "
    cString = cString & turn(cString) & "FILE1_10.MODELNO = " & MyParn(sModel)
    cString = cString & " GROUP BY FILE4_10.CODE , FILE4_10.DESCA ORDER BY CODE "
End If
Set data4.Recordset = myRecordSet(cString, con)
xcode.BoundText = sValue
If Not xcode.MatchedWithList Then xcode.BoundText = ""
End Sub

