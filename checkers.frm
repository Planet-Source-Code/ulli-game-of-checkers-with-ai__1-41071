VERSION 5.00
Begin VB.Form fCheckers 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Checkers"
   ClientHeight    =   9645
   ClientLeft      =   1395
   ClientTop       =   1770
   ClientWidth     =   11415
   ClipControls    =   0   'False
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Ausgefüllt
   Icon            =   "checkers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "checkers.frx":0152
   Moveable        =   0   'False
   ScaleHeight     =   643
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame fr 
      Caption         =   "Time to think"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Index           =   2
      Left            =   9135
      TabIndex        =   31
      Top             =   4305
      Width           =   1875
      Begin VB.HScrollBar scrTimeToThink 
         Height          =   240
         LargeChange     =   5
         Left            =   120
         Max             =   51
         Min             =   1
         TabIndex        =   32
         Top             =   375
         Value           =   5
         Width           =   1620
      End
      Begin VB.Label lbTime 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Approximate time to think"
         Top             =   690
         Width           =   1620
      End
   End
   Begin VB.CommandButton btEdit 
      Caption         =   "Edit Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10065
      TabIndex        =   1
      ToolTipText     =   "Set up individual piece positions"
      Top             =   810
      Width           =   930
   End
   Begin VB.PictureBox pcSquare 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   0
      Picture         =   "checkers.frx":045C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   29
      Top             =   1770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox pcSquare 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   0
      Picture         =   "checkers.frx":349E
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   28
      Top             =   810
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CheckBox ckView 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Show planned moves"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9120
      TabIndex        =   8
      ToolTipText     =   "Show the planned moves"
      Top             =   5685
      Width           =   1875
   End
   Begin VB.PictureBox picEinst 
      BackColor       =   &H00E0E0E0&
      Height          =   2790
      Left            =   9135
      Picture         =   "checkers.frx":64E0
      ScaleHeight     =   2730
      ScaleWidth      =   1800
      TabIndex        =   18
      ToolTipText     =   "Don't disturb me - I'm thinking"
      Top             =   5970
      Visible         =   0   'False
      Width           =   1860
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "thinking..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   345
         TabIndex        =   19
         Top             =   2415
         Width           =   990
      End
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   45
      Top             =   8940
   End
   Begin VB.CommandButton btNewGame 
      Caption         =   "Standard Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9135
      TabIndex        =   0
      ToolTipText     =   "Set up standard piece positions"
      Top             =   810
      Width           =   930
   End
   Begin VB.CommandButton btGo 
      BackColor       =   &H0080FF80&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9150
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Start a game"
      Top             =   8895
      Width           =   1845
   End
   Begin VB.Frame fr 
      Caption         =   "First Move"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   9135
      TabIndex        =   5
      Top             =   2880
      Width           =   1860
      Begin VB.OptionButton opComp 
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   6
         ToolTipText     =   "Computer makes first move"
         Top             =   360
         Width           =   990
      End
      Begin VB.OptionButton opHuman 
         Caption         =   "Human"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   7
         ToolTipText     =   "Human makes first move"
         Top             =   750
         Value           =   -1  'True
         Width           =   810
      End
   End
   Begin VB.Frame fr 
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   0
      Left            =   9135
      TabIndex        =   2
      Top             =   1455
      Width           =   1860
      Begin VB.OptionButton opPlayAlt 
         Caption         =   "Alternate Players"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   4
         ToolTipText     =   "Alternate players make the moves"
         Top             =   765
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton opPlaySelf 
         Caption         =   "Same Player"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   195
         TabIndex        =   3
         ToolTipText     =   "Same Player makes the moves for both sides"
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.ListBox lsPV 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   9150
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5970
      Width           =   1815
   End
   Begin VB.Label lblRules 
      BackStyle       =   0  'Transparent
      Caption         =   "International Rules"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2445
      TabIndex        =   30
      Top             =   585
      Width           =   1155
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BorderColor     =   &H00404040&
      Height          =   7710
      Index           =   3
      Left            =   945
      Top             =   945
      Width           =   7710
   End
   Begin VB.Label lblVN 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   540
      TabIndex        =   27
      Top             =   585
      Width           =   1155
   End
   Begin VB.Label lbPly 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
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
      Left            =   5235
      TabIndex        =   26
      ToolTipText     =   "Maximum ply analysed"
      Top             =   8895
      Width           =   945
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Max Ply"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4515
      TabIndex        =   25
      Top             =   8925
      Width           =   675
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoffs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   4575
      TabIndex        =   24
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label lbCutoff 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
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
      Left            =   5235
      TabIndex        =   23
      ToolTipText     =   "Number of Cutoffs"
      Top             =   9210
      Width           =   945
   End
   Begin VB.Label lbPosns 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
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
      Left            =   7845
      TabIndex        =   22
      ToolTipText     =   "Number of analysed positions"
      Top             =   9210
      Width           =   945
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posns visited"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   6660
      TabIndex        =   21
      Top             =   9240
      Width           =   1140
   End
   Begin VB.Label lbMsg 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3975
      TabIndex        =   13
      ToolTipText     =   "Various Messages"
      Top             =   225
      Width           =   6060
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Future Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   6705
      TabIndex        =   17
      Top             =   8925
      Width           =   1095
   End
   Begin VB.Label lbValue 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
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
      Left            =   7845
      TabIndex        =   16
      ToolTipText     =   "The best future value"
      Top             =   8895
      Width           =   945
   End
   Begin VB.Label lbYMM 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   705
      TabIndex        =   15
      Top             =   9105
      Width           =   1290
   End
   Begin VB.Label lbMoves 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2070
      TabIndex        =   14
      ToolTipText     =   "The Current Move"
      Top             =   9015
      Width           =   1650
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "to move"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2835
      TabIndex        =   12
      Top             =   270
      Width           =   690
   End
   Begin VB.Label lbSide 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Height          =   300
      Left            =   2460
      TabIndex        =   11
      ToolTipText     =   "Who's turn is it?"
      Top             =   225
      Width           =   300
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Checkers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   540
      TabIndex        =   10
      Top             =   150
      Width           =   1665
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BorderColor     =   &H00406040&
      BorderWidth     =   9
      Height          =   7860
      Index           =   0
      Left            =   870
      Top             =   870
      Width           =   7860
   End
End
Attribute VB_Name = "fCheckers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z
#Const Rules = 1 'International Version
'#Const Rules = 0 'AFC Version
Rem mark on
'History
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jun16,2000 Version 2.1
'Introduced Conditional Compilation:
'   Rules = 0       ACF-Rules (thanx to LimeyCliff)
'   Rules = 1       International Rules
'Changed board to marble to avoid Poor Rating - OK with you, Viper? ;=)
'Removed squarecolor - it is marble now
'Changed piece 3D color (don't need to Or it with the piece color)
'Added ulimited time to enable problem solution find
'Added board editing
'Put in more documentary notes
'Pretend to be a Screensaver - keep other Screensavers away
'Inhibited unload while thinking (thanx to Neph who reported this bug)
'Made MoveGen more symetrical by opposite direction stepping for B and W
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jun10,2000
'Increased the MoveList size and the PV size
'Added row shuffling
'Added small bonus to encourage jumping forward (effect only in intl version)
'Removed NumKings and NumMen (thanks - they helped me to find the overflow bug)
'Added Draw detection by looking for move loops in the PV
'Improved iterative deepening by breaking loop when winning or losing
'Put in some more documentary notes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jun7,2000
'Added square and piece highlighting (this created the effect that doubleclicking
'the board highlights all destination squares - dunno why it does that but its OK)
'Fixed Bug in promotion - was (semi?)-promoting while generating moves
'Moved squarecolor to enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jun6,2000
'Fixed Overflow Bug - was re-crowning a king on the 1st and 8th row
'Altered Search Opponent's Movelist to find ambiguous king moves
'Played with mousepointers
'Added icon
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jun2,2000
'Added a kind of highlight to the pieces to make'em look 3D
'Added NumKings and NumMen for Draw detection
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'May28,2000
'Moved most Constants into Enums
'Fixed some bugs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'May23,2000
'Hurray it plays
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Const SRCCOPY As Long = &HCC0020
Private Declare Sub SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long)
Private Const SPI_GETSCREENSAVEACTIVE As Long = 16
Private Const SPI_SETSCREENSAVEACTIVE As Long = 17

Private Enum Side
    White = 1
    Black = 2
    BothSides = 3
End Enum

Private Enum ColorScheme
    HiliteColor = vbYellow
    WhitePieceColor = vbRed
    BlackPieceColor = vbBlue
    White3DColor = &H8080FF
    Black3DColor = &HFF8080
    CrownColor = vbGreen            'or-ed to piece color
End Enum

Private Enum Diagonal
    NorthEast = -9
    NorthWest = -11
    SouthEast = 11
    SouthWest = 9
End Enum

Private Enum PieceValues           'and Bonusses / Penalties
    ManValue = 10
    KingValue = 30
    Maxmaterial = 12 * KingValue
    Infinity = 10000
    ForwardJumpBonus = ManValue / 10
End Enum

Private Type Board
    Value                    As Integer       'current board material balance
    SideToMove               As Byte
    Pieces(White To Black, _
           1 To 12)          As Byte          'lists of black and white pieces
    Squares(1 To 100)        As Byte          '8 by 8 plus guard squares (ForbiddenBit) on all four sides
    MovesListFrom            As Integer       'pointers into MoveList
    MovesListTo              As Integer
End Type

'The Squares are composed as follows    : (KingBit) (ForbiddenBit) (BlackBit) (WhiteBit) (four Bits pointer into PieceLists) - zero if free
'The Piecelists are composed as follows : (KingBit) (seven Bits position on board) - zero if captured
'Bytes are used because the Board is copied into TempBoard very frequently and should therefore be as short as possible
'The board itself is stretched to one dimension because a lone index is faster than two indexes

Private Const KingBit        As Byte = 128
Private Const ForbiddenBit   As Byte = 64
Private Const BlackBit       As Byte = 32
Private Const WhiteBit       As Byte = 16
Private Const FieldNumMask   As Byte = 127
Private Const PieceNumMask   As Byte = 15
Private Const DepthLimit     As Long = 48     'max ply searchable

Private x, y, i, j, k                         'common variables used for various purposes
Private SaverActive          As Long
Private LastX                As Single
Private LastY                As Single
Private Mode                 As String
Private GameEnds             As Boolean       'True when game ends
Private Terminate            As Boolean       'True when unloading requested
Private ClickEnabled         As Boolean
Private FormClicked          As Boolean
Private KeyDown              As Boolean
Private RequestCompleted     As Boolean
Private ClickButton          As Integer
Private Result               As Long          'The Search Result
Private ThisMove             As String
Private XN                   As String        'Decimal notation
Private HumanMove            As String
Private IterDepth            As Long          'Controls iterative deepening
Private TimeLimit            As Single
Private MaxPly               As Long          'Some statistical values
Private Cutoffs              As Long
Private PosnsVisited         As Currency
Private Forced               As Boolean       'True when forced move encountered
Private MoveDistance(1 To 4) As Long
Private MoveList(1 To DepthLimit * 12 * 13) As String 'Min size which will definitely not overflow
Private MoveListIx           As Long          'Index into MoveList
Private ResetIx              As Long          'Index is reset to this value if a longer MoveChain is found
Private Random(1 To 12)      As Long          'Used for row shuffeling
Private MaxChainLength       As Long
Private PV(0 To DepthLimit, _
           0 To DepthLimit)  As String        'Principal Variation
Private InPV                 As Long          '1 if in PV and 0 else
Private Board                As Board         'The Board

Private Sub Agent_RequestComplete(ByVal Request As Object)

    RequestCompleted = True

End Sub

Private Sub btEdit_Click()

    btGo.Enabled = False
    btEdit.Enabled = False
    btNewGame.Enabled = False
    GameEnds = True
    DrawBoard
    lbCutoff = ""
    lbMoves = ""
    lbPly = ""
    lbValue = ""
    lbSide.BackColor = WhitePieceColor
    With Board
        For i = White To Black
            For j = 1 To 12
                .Pieces(i, j) = 0
        Next j, i
        For i = 12 To 89
            .Squares(i) = 0
        Next i
        For i = 1 To 10 'fill guard Squares
            .Squares(i) = ForbiddenBit
            .Squares(i + 90) = ForbiddenBit
            .Squares(i * 10 - 9) = ForbiddenBit
            .Squares(i * 10) = ForbiddenBit
        Next i
        ClickEnabled = True
        lbMsg = "Press Spacebar when done"
        HumanMove = ""
        Do
            Do
                DoEvents
            Loop Until FormClicked Or KeyDown
            FormClicked = False
            i = Asc(HumanMove & Chr$(0))
            If i And (Int(i / 10) And 1) = ((i Mod 10) And 1) Then 'legal Square
                If ClickButton = vbRightButton Then 'Black
                    If .Squares(i) And (KingBit Or WhiteBit) Then
                        .Squares(i) = 0
                        ErasePiece (i)
                      ElseIf .Squares(i) = 0 Then 'NOT .SQUARES(I)...
                        If i > 81 Then
                            .Squares(i) = BlackBit Or KingBit
                            DrawBlackKing (i)
                          Else 'NOT I...
                            .Squares(i) = BlackBit
                            DrawBlackMan (i)
                        End If
                      ElseIf (.Squares(i) And KingBit) = 0 Then 'NOT .SQUARES(I)...
                        .Squares(i) = .Squares(i) Or KingBit
                        DrawBlackKing (i)
                    End If
                  Else                              'White 'NOT CLICKBUTTON...
                    If .Squares(i) And (KingBit Or BlackBit) Then
                        .Squares(i) = 0
                        ErasePiece (i)
                      ElseIf .Squares(i) = 0 Then 'NOT .SQUARES(I)...
                        If i < 21 Then
                            .Squares(i) = WhiteBit Or KingBit
                            DrawWhiteKing (i)
                          Else 'NOT I...
                            .Squares(i) = WhiteBit
                            DrawWhiteMan (i)
                        End If
                      ElseIf (.Squares(i) And KingBit) = 0 Then 'NOT .SQUARES(I)...
                        .Squares(i) = .Squares(i) Or KingBit
                        DrawWhiteKing (i)
                    End If
                End If
            End If
            HumanMove = ""
            lbMoves = ""
            ClickButton = 0
        Loop Until KeyDown
        j = 0
        k = 0
        .Value = 0
        For i = 89 To 12 Step -1
            If .Squares(i) And WhiteBit Then
                j = j + 1
                If j > 12 Then
                    .Squares(i) = 0
                    ErasePiece i
                  Else 'NOT J...
                    .Pieces(White, j) = i Or (KingBit And .Squares(i))
                    .Squares(i) = .Squares(i) Or j
                    If .Squares(i) And KingBit Then
                        .Value = .Value + KingValue
                      Else 'NOT .SQUARES(I)...
                        .Value = .Value + ManValue
                    End If
                End If
            End If
        Next i
        For i = 12 To 89
            If .Squares(i) And BlackBit Then
                k = k + 1
                If k > 12 Then
                    .Squares(i) = 0
                    ErasePiece i
                  Else 'NOT K...
                    .Pieces(Black, k) = i Or (KingBit And .Squares(i))
                    .Squares(i) = .Squares(i) Or k
                    If .Squares(i) And KingBit Then
                        .Value = .Value - KingValue
                      Else 'NOT .SQUARES(I)...
                        .Value = .Value - ManValue
                    End If
                End If
            End If
        Next i
    End With 'BOARD
    btGo.Enabled = True
    btEdit.Enabled = True
    btNewGame.Enabled = True
    fr(0).Enabled = True
    fr(1).Enabled = True
    ClickEnabled = False
    KeyDown = False

End Sub

Private Sub btGo_Click()

    fr(0).Enabled = False
    fr(1).Enabled = False
    Board.SideToMove = White
    btGo.Enabled = False
    PlayGame
    fr(0).Enabled = True
    fr(1).Enabled = True

End Sub

Private Sub btNewGame_Click()

    GameEnds = True
    Form_Load
    Form_Paint
    btGo.Enabled = True
    btEdit.Enabled = True
    fr(0).Enabled = True
    fr(1).Enabled = True

End Sub

Private Sub ckView_Click()

    lsPV.Clear
    If ckView Then
        lsPV.ToolTipText = "Here are the planned moves"
      Else 'CKVIEW = 0'CKVIEW = FALSE
        lsPV.ToolTipText = "No moves showing"
    End If

End Sub

Private Sub ComputerMoves()

  Dim CMResult

    With Board
        picEinst.Visible = True
        btNewGame.Enabled = False
        btEdit.Enabled = False
        DoEvents
        TimeLimit = Timer + IIf(scrTimeToThink > 50, 1E+20, scrTimeToThink * 3)
        IterDepth = 0
        InPV = 0
        PosnsVisited = 0
        MaxPly = 0
        Cutoffs = 0
        Forced = False
        lbMoves.ForeColor = &H404040
        lbYMM = "Thinking about"
        MousePointer = vbHourglass
        Do 'iterative search deepening
            .MovesListFrom = 1
            .MovesListTo = 0
            IterDepth = IterDepth + 1
            CMResult = Search(Board, 0, IterDepth, -Infinity, Infinity)
            InPV = 1
            If ckView = vbChecked Then
                lbMoves = FieldToXN(Asc(PV(0, 0))) & "-" & FieldToXN(Asc(Mid$(PV(0, 0), 2)))
              Else 'NOT CKVIEW...
                lbMoves = "??-??"
            End If
            DoEvents
        Loop While Timer < TimeLimit And IterDepth < DepthLimit And Not Forced And Abs(CMResult) <= Maxmaterial
        MousePointer = vbNormal
        picEinst.Visible = (Mode = "CC")
        lbPly = MaxPly
        lbCutoff = Cutoffs
        lbPosns = PosnsVisited
        If Forced Then
            lbValue = "Unknown"
          Else 'FORCED = 0'FORCED = FALSE
            lbValue = CMResult
        End If
        lbValue.ForeColor = IIf(.SideToMove = White, WhitePieceColor, BlackPieceColor)
        Select Case CMResult
          Case -Infinity
            lbMsg = "I have lost."
            GameEnds = True
          Case Is < -Maxmaterial
            lbMsg = "I resign."
            GameEnds = True
          Case Else
            If CMResult > Maxmaterial Then
                i = Infinity - CMResult
                lbMsg = "I win in " & i & " move" & IIf(i > 1, "s.", ".")
            End If
            lbYMM = "My move"
            lbMoves.ForeColor = lbValue.ForeColor
            lbMoves = FieldToXN(Asc(PV(0, 0))) & "-" & FieldToXN(Asc(Mid$(PV(0, 0), 2)))
            If ckView = vbChecked Then
                lsPV.Clear
                If Not Forced Then
                    lsPV.AddItem "PLANNED MOVES"
                    i = 1
                    Do While PV(0, i) <> ""
                        If PV(0, i + 1) = "" Then
                            lsPV.AddItem "    " & FieldToXN(Asc(PV(0, i))) & "-" & FieldToXN(Asc(Mid$(PV(0, i), 2)))
                            Exit Do '>---> Loop
                          Else 'NOT PV(0,...
                            lsPV.AddItem "    " & FieldToXN(Asc(PV(0, i))) & "-" & FieldToXN(Asc(Mid$(PV(0, i), 2))) & vbTab & _
                                         FieldToXN(Asc(PV(0, i + 1))) & "-" & FieldToXN(Asc(Mid$(PV(0, i + 1), 2)))
                        End If
                        i = i + 2
                    Loop
                End If
            End If
            HiliteSquare Asc(PV(0, 0))
            For i = 2 To Len(PV(0, 0)) Step 3
                k = Asc(Mid$(PV(0, 0), i, 1))
                HiliteSquare k
            Next i
            For i = i To 0 Step -1
                DoEvents
                Sleep 200
            Next i
            MakeMove Board, PV(0, 0), True, True
            UnHiliteSquare k
            If IterDepth > 15 Then
                If PV(0, IterDepth - 4) = Mid$(PV(0, IterDepth - 2), 2) & Left$(PV(0, IterDepth - 2), 1) And _
                   PV(0, IterDepth - 3) = Mid$(PV(0, IterDepth - 1), 2) & Left$(PV(0, IterDepth - 1), 1) Then 'looping
                    If Mode = "CC" Then
                        lbMsg = "Neither side can win any more"
                        GameEnds = True
                      Else 'NOT MODE...
                        GameEnds = (MsgBox("This game looks like a Draw - continue anyway?", vbYesNo Or vbQuestion) = vbNo)
                    End If
                End If
            End If
            .SideToMove = BothSides - .SideToMove 'toggle side to move
        End Select
        DoEvents
    End With 'BOARD

End Sub

Private Sub DrawBlackKing(ByVal Square As Long)

    Square = Square - 1
    x = (Square Mod 10) * 64 + 29
    y = (Square \ 10) * 64 + 29
    FillColor = BlackPieceColor Or Black3DColor
    Circle (x, y), 24, BlackPieceColor Or Black3DColor
    x = x + 3
    y = y + 3
    FillColor = BlackPieceColor
    Circle (x, y), 25, BlackPieceColor
    FillColor = BlackPieceColor Or CrownColor
    Circle (x, y), 12, BlackPieceColor Or CrownColor

End Sub

Private Sub DrawBlackMan(ByVal Square As Long)

    Square = Square - 1
    FillColor = Black3DColor
    Circle ((Square Mod 10) * 64 + 30, (Square \ 10) * 64 + 29), 24, Black3DColor
    FillColor = BlackPieceColor
    Circle ((Square Mod 10) * 64 + 32, (Square \ 10) * 64 + 32), 25, BlackPieceColor

End Sub

Private Sub DrawBoard()

    For x = 64 To 575 Step 64
        For y = 64 To 575 Step 64
            BitBlt Me.hDC, x, y, 64, 64, pcSquare(IIf((x And 64) Xor (y And 64), 0, 1)).hDC, 0, 0, SRCCOPY
    Next y, x

End Sub

Private Sub DrawWhiteKing(ByVal Square As Long)

    Square = Square - 1
    x = (Square Mod 10) * 64 + 29
    y = (Square \ 10) * 64 + 29
    FillColor = WhitePieceColor Or White3DColor
    Circle (x, y), 24, WhitePieceColor Or White3DColor
    x = x + 3
    y = y + 3
    FillColor = WhitePieceColor
    Circle (x, y), 25, WhitePieceColor
    FillColor = WhitePieceColor Or CrownColor
    Circle (x, y), 12, WhitePieceColor Or CrownColor

End Sub

Private Sub DrawWhiteMan(ByVal Square As Long)

    Square = Square - 1
    FillColor = White3DColor
    Circle ((Square Mod 10) * 64 + 30, (Square \ 10) * 64 + 29), 24, White3DColor
    FillColor = WhitePieceColor
    Circle ((Square Mod 10) * 64 + 32, (Square \ 10) * 64 + 32), 25, WhitePieceColor

End Sub

Private Sub ErasePiece(ByVal Square As Long)

    Square = Square - 1
    x = (Square Mod 10) * 64
    y = (Square \ 10) * 64
    BitBlt Me.hDC, x, y, 64, 64, pcSquare(0).hDC, 0, 0, SRCCOPY

End Sub

Private Function Evaluate(Board As Board) As Long

  'plays good enough as it is - no positonal evaluation necessary

    With Board
        If .SideToMove = White Then
            Evaluate = .Value
          Else 'NOT .SIDETOMOVE...
            Evaluate = -.Value
        End If
    End With 'BOARD

End Function

Private Sub FieldToLbl(Square As Long)

    If Len(lbMoves) > 1 Then
        lbMoves = lbMoves & "-"
    End If
    HumanMove = HumanMove & Chr$(Square)
    lbMoves = lbMoves & FieldToXN(Square)

End Sub

Private Function FieldToXN(Square As Long) As String

    XN = Format$(Square, "00")
    FieldToXN = Chr$(Asc(Mid$(XN, 2)) + 15) & Chr$(Asc("9") - Val(Left$(XN, 1)))

End Function

Private Sub Form_Initialize()

  'Pretend to be a Screensaver

    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0&, SaverActive, 0&
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0&, ByVal 0&, 0&

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown = True
    KeyCode = 0       'Consume

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyDown = False
    KeyCode = 0

End Sub

Private Sub Form_Load()

    fr(0).BackColor = BackColor
    fr(1).BackColor = BackColor
    ckView.BackColor = BackColor
    opPlaySelf.BackColor = BackColor
    opPlayAlt.BackColor = BackColor
    opComp.BackColor = BackColor
    opHuman.BackColor = BackColor
    With Board
        For i = 12 To 89
            .Squares(i) = 0
        Next i
        For i = 1 To 10 'fill guard Squares
            .Squares(i) = ForbiddenBit
            .Squares(i + 90) = ForbiddenBit
            .Squares(i * 10 - 9) = ForbiddenBit
            .Squares(i * 10) = ForbiddenBit
        Next i
        .Pieces(White, 1) = 62        'Place pieces on board and in piecelists
        .Squares(62) = WhiteBit Or 1
        .Pieces(White, 2) = 64
        .Squares(64) = WhiteBit Or 2
        .Pieces(White, 3) = 66
        .Squares(66) = WhiteBit Or 3
        .Pieces(White, 4) = 68
        .Squares(68) = WhiteBit Or 4
        .Pieces(White, 5) = 73
        .Squares(73) = WhiteBit Or 5
        .Pieces(White, 6) = 75
        .Squares(75) = WhiteBit Or 6
        .Pieces(White, 7) = 77
        .Squares(77) = WhiteBit Or 7
        .Pieces(White, 8) = 79
        .Squares(79) = WhiteBit Or 8
        .Pieces(White, 9) = 82
        .Squares(82) = WhiteBit Or 9
        .Pieces(White, 10) = 84
        .Squares(84) = WhiteBit Or 10
        .Pieces(White, 11) = 86
        .Squares(86) = WhiteBit Or 11
        .Pieces(White, 12) = 88
        .Squares(88) = WhiteBit Or 12

        .Pieces(Black, 1) = 33
        .Squares(33) = BlackBit Or 1
        .Pieces(Black, 2) = 35
        .Squares(35) = BlackBit Or 2
        .Pieces(Black, 3) = 37
        .Squares(37) = BlackBit Or 3
        .Pieces(Black, 4) = 39
        .Squares(39) = BlackBit Or 4
        .Pieces(Black, 5) = 22
        .Squares(22) = BlackBit Or 5
        .Pieces(Black, 6) = 24
        .Squares(24) = BlackBit Or 6
        .Pieces(Black, 7) = 26
        .Squares(26) = BlackBit Or 7
        .Pieces(Black, 8) = 28
        .Squares(28) = BlackBit Or 8
        .Pieces(Black, 9) = 13
        .Squares(13) = BlackBit Or 9
        .Pieces(Black, 10) = 15
        .Squares(15) = BlackBit Or 10
        .Pieces(Black, 11) = 17
        .Squares(17) = BlackBit Or 11
        .Pieces(Black, 12) = 19
        .Squares(19) = BlackBit Or 12
        .Value = 0
        .SideToMove = White
    End With 'BOARD
    For i = 1 To 12
        Random(i) = 0
    Next i
    Randomize
    'shuffle the three rows
    For i = 1 To 4
        Do
            j = Int(Rnd * 4) + 1
            If Random(j) = 0 Then
                Random(j) = i
                Exit Do '>---> Loop
            End If
        Loop
    Next i
    For i = 5 To 8
        Do
            j = Int(Rnd * 4) + 5
            If Random(j) = 0 Then
                Random(j) = i
                Exit Do '>---> Loop
            End If
        Loop
    Next i
    For i = 9 To 12
        Do
            j = Int(Rnd * 4) + 9
            If Random(j) = 0 Then
                Random(j) = i
                Exit Do '>---> Loop
            End If
        Loop
    Next i
    lbSide.BackColor = WhitePieceColor
    lbValue = ""
    lbPly = ""
    lbCutoff = ""
    lbPosns = ""
    lbMoves = ""
    MoveDistance(1) = NorthWest   '5-1 = 4 = SE (opposite direction is forbidden during captures)
    MoveDistance(2) = NorthEast   '5-2 = 3 = SW
    MoveDistance(3) = SouthWest   '5-3 = 2 = NE
    MoveDistance(4) = SouthEast   '5-4 = 1 = NW
    ClickEnabled = False
    scrTimeToThink = 5            '5/10 Minutes = 30 seconds
    opComp_Click
    ckView = vbUnchecked
    ckView_Click
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules = 0 Then
    lblRules = "ACF Rules"
#End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    lblVN = "Version " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If x > 63 And x < 576 And y > 64 And y < 576 Then
        If ClickEnabled Then
            FieldToLbl Int(x / 64) + 11 + (Int(y / 64) - 1) * 10
        End If
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If ClickEnabled Then
        If x > 63 And x < 576 And y > 63 And y < 576 Then
            MousePointer = vbCustom
          Else 'NOT X...
            MousePointer = vbNormal
        End If
      Else 'CLICKENABLED = 0'CLICKENABLED = FALSE
        MousePointer = IIf(btNewGame.Enabled, vbNormal, vbHourglass)
    End If
    LastX = x
    LastY = y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    FormClicked = ClickEnabled And x > 63 And x < 576 And y > 63 And y < 576
    ClickButton = Button

End Sub

Private Sub Form_Paint()

    DrawBoard
    For i = 1 To 12
        j = Board.Pieces(Black, i)
        If j And KingBit Then
            DrawBlackKing j And FieldNumMask
          Else 'NOT J...
            DrawBlackMan j And FieldNumMask
        End If
        j = Board.Pieces(White, i)
        If j And KingBit Then
            DrawWhiteKing j And FieldNumMask
          Else 'NOT J...
            DrawWhiteMan j And FieldNumMask
        End If
    Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = Not (ClickEnabled Or btGo.Enabled Or GameEnds)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Terminate = True
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, SaverActive, ByVal 0&, 0&
    End

End Sub

Private Sub GenerateMoves(Board As Board, Square As Long, Depth As Long, MovesSoFar As String, ForbiddenDirection As Long)

  Dim TempBoard     As Board
  Dim Direction     As Long
  Dim Slide         As Long
  Dim Adjacent      As Long
  Dim Beyond        As Long
  Dim Captured      As String

    With Board
        If .SideToMove = White Then
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules <> 0 Then 'International Version - King moves any number of squares
            If .Squares(Square) And KingBit Then     'white King
                For Direction = 1 To 4
                    If Direction <> ForbiddenDirection Then
                        Adjacent = MoveDistance(Direction)
                        Captured = ""
                        Slide = Square
                        Do Until .Squares(Slide + Adjacent) And (ForbiddenBit Or WhiteBit)
                            Slide = Slide + Adjacent  'sliding a king
                            If .Squares(Slide) = 0 Then
                                If Captured = "" Then 'nothing captured yet
                                    If Depth = 0 Then
                                        RecordMove Board, Chr$(Square) & Chr$(Slide)
                                    End If
                                  Else 'has found a pice to capture and there is an empty Square beyond that 'NOT CAPTURED...
                                    ThisMove = Chr$(Square) & Chr$(Slide) & Captured
                                    TempBoard = Board
                                    MakeMove TempBoard, ThisMove, False, False 'land on empty Square and see if there are more captures?
                                    GenerateMoves TempBoard, Slide, Depth + 1, MovesSoFar & ThisMove, 5 - Direction
                                End If
                              Else 'NOT .SQUARES(SLIDE)...
                                If Captured = "" Then 'nothing captured yet
                                    If .Squares(Slide) And BlackBit Then 'found a piece to capture
                                        Captured = Chr$(Slide)
                                    End If
                                  Else 'NOT CAPTURED...
                                    Exit Do 'do not capture two pieces in one slide '>---> Loop
                                End If
                            End If
                        Loop
                    End If
                Next Direction
              Else                                 'white Man 'NOT .SQUARES(SQUARE)...
#End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules = 0 Then 'ACF Version - men jumping fwd only, kings in all four directions
                For Direction = 1 To IIf(.Squares(Square) And KingBit, 4, 2)
#Else          'International Version - jumping in all four directions
                    For Direction = 1 To 4
                        If Direction <> ForbiddenDirection Then
#End If
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            Adjacent = Square + MoveDistance(Direction)
                            Beyond = Adjacent + MoveDistance(Direction)
                            If .Squares(Adjacent) And BlackBit Then
                                If .Squares(Beyond) = 0 Then 'this is a capture - any more captures possible?
                                    ThisMove = Chr$(Square) & Chr$(Beyond) & Chr$(Adjacent)
                                    TempBoard = Board
                                    MakeMove TempBoard, ThisMove, False, False
                                    GenerateMoves TempBoard, Beyond, Depth + 1, MovesSoFar & ThisMove, 5 - Direction
                                End If
                              Else 'NOT .SQUARES(ADJACENT)...
                                If .Squares(Adjacent) = 0 And (Adjacent < Square Or .Squares(Square) And KingBit) And Depth = 0 Then 'no capture
                                    RecordMove Board, Chr$(Square) & Chr$(Adjacent)
                                End If
                            End If
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules <> 0 Then 'International Version
                        End If
                    Next Direction
                End If
#Else
            Next Direction
#End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          Else 'NOT .SIDETOMOVE...
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules <> 0 Then 'International Version - King moves any number of squares
            If .Squares(Square) And KingBit Then     'black King
                For Direction = 4 To 1 Step -1
                    If Direction <> ForbiddenDirection Then
                        Adjacent = MoveDistance(Direction)
                        Captured = ""
                        Slide = Square
                        Do Until .Squares(Slide + Adjacent) And (ForbiddenBit Or BlackBit)
                            Slide = Slide + Adjacent
                            If .Squares(Slide) = 0 Then
                                If Captured = "" Then
                                    If Depth = 0 Then
                                        RecordMove Board, Chr$(Square) & Chr$(Slide)
                                    End If
                                  Else 'NOT CAPTURED...
                                    ThisMove = Chr$(Square) & Chr$(Slide) & Captured
                                    TempBoard = Board
                                    MakeMove TempBoard, ThisMove, False, False
                                    GenerateMoves TempBoard, Slide, Depth + 1, MovesSoFar & ThisMove, 5 - Direction
                                End If
                              Else 'NOT .SQUARES(SLIDE)...
                                If Captured = "" Then
                                    If .Squares(Slide) And WhiteBit Then
                                        Captured = Chr$(Slide)
                                    End If
                                  Else 'NOT CAPTURED...
                                    Exit Do '>---> Loop
                                End If
                            End If
                        Loop
                    End If
                Next Direction
              Else                                 'black Man 'NOT .SQUARES(SQUARE)...
#End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules = 0 Then   'ACF Version - men jumping fwd only, kings in all four directions
                For Direction = 4 To IIf(.Squares(Square) And KingBit, 1, 3) Step -1
#Else          'International Version - jumping in all four directions
                    For Direction = 4 To 1 Step -1
                        If Direction <> ForbiddenDirection Then
#End If
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            Adjacent = Square + MoveDistance(Direction)
                            Beyond = Adjacent + MoveDistance(Direction)
                            If .Squares(Adjacent) And WhiteBit Then
                                If .Squares(Beyond) = 0 Then
                                    ThisMove = Chr$(Square) & Chr$(Beyond) & Chr$(Adjacent)
                                    TempBoard = Board
                                    MakeMove TempBoard, ThisMove, False, False
                                    GenerateMoves TempBoard, Beyond, Depth + 1, MovesSoFar & ThisMove, 5 - Direction
                                End If
                              Else 'NOT .SQUARES(ADJACENT)...
                                If .Squares(Adjacent) = 0 And (Adjacent > Square Or .Squares(Square) And KingBit) And Depth = 0 Then  'no capture
                                    RecordMove Board, Chr$(Square) & Chr$(Adjacent)
                                End If
                            End If
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules <> 0 Then 'International Version
                        End If
                    Next Direction
                End If
#Else
            Next Direction
#End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        RecordMove Board, MovesSoFar
    End With 'BOARD

End Sub

Private Sub HilitePiece(ByVal Square As Long)

    Square = Square - 1
    DrawStyle = vbDot
    FillStyle = vbFSTransparent
    Circle ((Square Mod 10) * 64 + 32, (Square \ 10) * 64 + 32), 25, HiliteColor
    DrawStyle = vbSolid
    FillStyle = vbFSSolid

End Sub

Private Sub HiliteSquare(ByVal Square As Long)

    Square = Square - 1
    x = (Square Mod 10) * 64 + 1
    y = (Square \ 10) * 64 + 1
    DrawStyle = vbDot
    Line (x, y)-Step(61, 0), HiliteColor
    Line -Step(0, 61), HiliteColor
    Line -Step(-61, 0), HiliteColor
    Line -Step(0, -61), HiliteColor
    DrawStyle = vbSolid

End Sub

Private Sub HumanMoves()

  Dim FieldNum As Long
  Dim Hilited As String

    With Board
        MaxChainLength = 1
        .MovesListFrom = 1
        ResetIx = 0
        InPV = 0
        For i = 1 To 12
            FieldNum = .Pieces(.SideToMove, i) And FieldNumMask
            If FieldNum Then
                GenerateMoves Board, FieldNum, 0, "", 0 'generate moves for opponent
            End If
        Next i
        DoEvents
        If MaxChainLength >= 2 Then 'opponent can move
            lbMoves.ForeColor = lbSide.BackColor
            lbYMM = "Your move"
            lbMoves = ""
            If MoveListIx = 1 Then
                MousePointer = vbHourglass
                lbMsg = "You have only one move; it will be made for you."
                FieldToLbl Asc(MoveList(1))
                FieldToLbl Asc(Mid$(MoveList(1), 2))
                HiliteSquare Asc(MoveList(1))
                For i = 2 To Len(MoveList(1)) Step 3
                    k = Asc(Mid$(MoveList(1), i, 1))
                    HiliteSquare k
                Next i
                Do
                    DoEvents
                Loop While Len(lbMsg)
                MakeMove Board, MoveList(1), True, True
                UnHiliteSquare k
                Form_MouseMove 0, 0, LastX, LastY 'restore mousepointer
                MaxChainLength = 1
              Else 'NOT MOVELISTIX...
                ClickEnabled = True
                Form_MouseMove 0, 0, LastX, LastY
                lbMsg = "Enter your move:"
                HumanMove = ""
                For i = 2 To MoveListIx
                    If Left$(MoveList(i), 1) <> Left$(MoveList(1), 1) Then
                        Exit For '>---> Next
                    End If
                Next i
                If i > MoveListIx Then           'only one piece can move
                    HilitePiece Asc(MoveList(1))
                End If
                Do
                    FormClicked = False
                    btNewGame.Enabled = True
                    btEdit.Enabled = True
                    Do
                        DoEvents
                        If Terminate Then
                            End
                        End If
                    Loop While btNewGame.Enabled And Not FormClicked
                    If FormClicked Then 'search opponent's move list
                        For j = 1 To Len(Hilited)
                            UnHiliteSquare Asc(Mid$(Hilited, j, 1))
                        Next j
                        Hilited = Left$(HumanMove, 1)
                        j = 0
                        For i = 1 To MoveListIx
                            For x = 1 To Len(HumanMove)
                                If Mid$(HumanMove, x, 1) <> Mid$(MoveList(i), x, 1) And Mid$(HumanMove, x, 1) <> "?" Then 'this is not it
                                    Exit For '>---> Next
                                End If
                            Next x
                            If x > Len(HumanMove) Then
                                j = j + 1
                                k = i
                                For x = 2 To Len(HumanMove) + 3 Step 3
                                    Hilited = Hilited & Mid$(MoveList(i), x, 1)
                                Next x
                            End If
                        Next i
                        Select Case j
                          Case 0        'not in opponent's movelist
                            lbMsg = "You entered an illegal move."
                            HumanMove = ""
                            lbMoves = ""
                          Case 1        'move uniquely defined - make it
                            lbMsg = ""
                            MakeMove Board, MoveList(k), True, True
                            For j = 1 To Len(Hilited)
                                UnHiliteSquare Asc(Mid$(Hilited, j, 1))
                            Next j
                            Exit Do     'and exit '>---> Loop
                          Case Else     'not unique yet
                            For j = 1 To Len(Hilited)
                                HiliteSquare Asc(Mid$(Hilited, j, 1))
                            Next j
                            If Len(HumanMove) >= 2 Then
                                HumanMove = HumanMove & "?" & Right$(HumanMove, 1)
                                lbMsg = "More..."
                            End If
                        End Select
                    End If
                Loop
                ClickEnabled = False
                If Rnd < 0.1 Then
                    lbMsg = "Do you mind me smoking?"
                End If
            End If
          Else 'opponent cannot move 'NOT MAXCHAINLENGTH...
            lbMsg = "You have lost."
            GameEnds = True
        End If
        .SideToMove = BothSides - .SideToMove 'toggle side to move
    End With 'BOARD
    DoEvents

End Sub

Private Sub lbMsg_Change()

    tmr.Enabled = False 'restart timer
    tmr.Enabled = True

End Sub

Private Sub MakeMove(Board As Board, ByVal MoveChain As String, ForReal As Boolean, Paint As Boolean)

  Dim FromField As Long
  Dim ToField   As Long
  Dim Captured  As Long

    With Board
        Do
            FromField = Asc(MoveChain)
            ToField = Asc(Mid$(MoveChain, 2))
            If Len(MoveChain) > 2 Then
                Captured = Asc(Mid$(MoveChain, 3))
            End If
            .Squares(ToField) = .Squares(FromField)
            If Len(MoveChain) < 4 And ForReal Then 'this is a real move (not while genereating) and it is the last (or only) of a chain
                If (.SideToMove = White And ToField < 21) Or (.SideToMove = Black And ToField > 81) Then
                    If (.Squares(FromField) And KingBit) = 0 Then   'Crown a new king
                        .Squares(ToField) = .Squares(ToField) Or KingBit
                        If .SideToMove = White Then 'White promotes
                            .Value = .Value - ManValue + KingValue 'correct value for promotion
                          Else                      'Black promotes 'NOT .SIDETOMOVE...
                            .Value = .Value + ManValue - KingValue
                        End If
                    End If
                End If
            End If
            .Squares(FromField) = 0
            .Pieces(.SideToMove, .Squares(ToField) And PieceNumMask) = ToField Or (KingBit And .Squares(ToField))
            If Captured Then
                If .SideToMove = White Then          'White captures Black
                    If .Squares(Captured) And KingBit Then
                        .Value = .Value + KingValue  'correct value for capture
                      Else 'NOT .SQUARES(CAPTURED)...
                        .Value = .Value + ManValue
                    End If
                    If ForReal Then
                        If FromField > ToField Then  'Bonus for white
                            .Value = .Value + ForwardJumpBonus
                          Else                       'Bonus for black 'NOT FROMFIELD...
                            .Value = .Value - ForwardJumpBonus
                        End If
                    End If
                  Else                                'Black captures White 'NOT .SIDETOMOVE...
                    If .Squares(Captured) And KingBit Then
                        .Value = .Value - KingValue   'correct value for capture
                      Else 'NOT .SQUARES(CAPTURED)...
                        .Value = .Value - ManValue
                    End If
                    If ForReal Then
                        If FromField < ToField Then  'Bonus for black
                            .Value = .Value - ForwardJumpBonus
                          Else                       'Bonus for white 'NOT FROMFIELD...
                            .Value = .Value + ForwardJumpBonus
                        End If
                    End If
                End If
                .Pieces(BothSides - .SideToMove, .Squares(Captured) And PieceNumMask) = 0
                .Squares(Captured) = IIf(ForReal, 0, ForbiddenBit) 'do not capture this piece again (while generating)
            End If
            If Paint Then               'update display
                ErasePiece FromField
                Sleep 50
                If .SideToMove = White Then
                    If .Squares(ToField) And KingBit Then
                        DrawWhiteKing ToField
                      Else 'NOT .SQUARES(TOFIELD)...
                        DrawWhiteMan ToField
                    End If
                  Else 'NOT .SIDETOMOVE...
                    If .Squares(ToField) And KingBit Then
                        DrawBlackKing ToField
                      Else 'NOT .SQUARES(TOFIELD)...
                        DrawBlackMan ToField
                    End If
                End If
                If Captured Then
                    ErasePiece Captured
                End If
                Sleep 700
            End If
            MoveChain = Mid$(MoveChain, 4) 'next move in movechain
        Loop While Len(MoveChain) >= 3    'none left - exit
    End With 'BOARD

End Sub

Private Sub opComp_Click()

    If opComp Then
        If opPlaySelf Then
            Mode = "CC"
          Else 'OPPLAYSELF = 0'OPPLAYSELF = FALSE
            Mode = "CH"
        End If
      Else 'OPCOMP = 0'OPCOMP = FALSE
        If opPlaySelf Then
            Mode = "HH"
          Else 'OPPLAYSELF = 0'OPPLAYSELF = FALSE
            Mode = "HC"
        End If
    End If

End Sub

Private Sub opHuman_Click()

    opComp_Click

End Sub

Private Sub opPlayAlt_Click()

    opComp_Click

End Sub

Private Sub opPlaySelf_Click()

    opComp_Click

End Sub

Private Sub PlayGame()

    With Board
        GameEnds = False
        Do Until GameEnds
            Select Case Mode
              Case "CC"
                ComputerMoves
              Case "HH"
                HumanMoves
              Case "CH"
                If .SideToMove = White Then
                    ComputerMoves
                  Else 'NOT .SIDETOMOVE...
                    HumanMoves
                End If
              Case "HC"
                If .SideToMove = White Then
                    HumanMoves
                  Else 'NOT .SIDETOMOVE...
                    ComputerMoves
                End If
            End Select
            lbSide.BackColor = IIf(.SideToMove = White, WhitePieceColor, BlackPieceColor)
        Loop
    End With 'BOARD
    'game finished
    lbYMM = ""
    lbMoves.ForeColor = vbBlack
    lbMoves = "Game ends"
    lbSide.BackColor = BackColor
    btNewGame.Enabled = True
    btEdit.Enabled = True
    picEinst.Visible = False
    MousePointer = vbNormal

End Sub

Private Sub RecordMove(Board As Board, MoveChain As String)

  'minimal move recording - we might not need all the moves because of cutoffs
  'a move consists of 2 or 3 bytes (or n*3 bytes if it is a capture series)
  'asc(FirstByte) is from-Square ; asc(SecondByte) is to-Square ; asc(ThirdByte) is captured Square if applicable
  'all moves go into the same list which is subdivided into segments; one for each ply; the segment boundaries
  'are stored with the board

    With Board
        k = Len(MoveChain)
        If k > MaxChainLength Then                      'found a longer MoveChain; reset index to start of segment
            MaxChainLength = k
            MoveListIx = ResetIx
        End If
        If k = MaxChainLength Then                      'this move has same length as all others previously recorded
            If MoveChain <> MoveList(.MovesListFrom) Or InPV = 0 Then
                MoveListIx = MoveListIx + 1
                MoveList(MoveListIx) = MoveChain
            End If
        End If
    End With 'BOARD

End Sub

Private Sub scrTimeToThink_Change()

    lbTime.BackColor = &HE0E0E0
    lbTime.ForeColor = vbBlack
    Select Case scrTimeToThink
      Case Is < 10
        lbTime = scrTimeToThink * 6 & " Seconds"
      Case 10
        lbTime = "1 Minute"
      Case 51
        lbTime = "Unlimited"
        lbTime.BackColor = vbYellow
        lbTime.ForeColor = vbRed
      Case Else
        lbTime = scrTimeToThink / 10 & " Minutes"
    End Select

End Sub

Private Sub scrTimeToThink_Scroll()

    scrTimeToThink_Change

End Sub

Private Function Search(Board As Board, Depth, MinDepth, Alpha, Beta) As Long

  Dim TempBoard     As Board
  Dim Square        As Long
  Dim BestResult    As Long
  Dim PcMvNum       As Long
  Dim CurrMoveChain As String

    If Depth > MaxPly Then
        MaxPly = Depth
    End If
    With Board
        PosnsVisited = PosnsVisited + 1
        If InPV Then
            If Len(PV(0, Depth)) Then
                MoveList(.MovesListFrom) = PV(0, Depth)  'analyse the most promising move first
              Else                                       'this ensures that a: the Principal Variation is established as soon as possible or 'NOT LEN(PV(0,...
                InPV = 0                                 '                  b: a cutoff is triggered as early as possible
            End If
        End If
        MaxChainLength = 1
        MoveListIx = .MovesListFrom - 1 + InPV           'dont touch first entry while in pv
        ResetIx = MoveListIx
        For PcMvNum = 1 To 12
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If Rules = 0 Then 'AFC Rules - Choose jump
            If MaxChainLength >= 3 Then
                MaxChainLength = 3
                ResetIx = MoveListIx
            End If
#End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Square = .Pieces(.SideToMove, Random(PcMvNum)) And FieldNumMask
            If Square Then
                GenerateMoves Board, Square, 0, "", 0    'generate all legal moves for each live piece of side to move
            End If
        Next PcMvNum
        PV(Depth, Depth) = ""                            'clear Principal Variation
        .MovesListTo = MoveListIx
        If (MaxChainLength = 2 And Depth >= MinDepth) Or (Depth = DepthLimit) Then 'quiescence found and min depth searched or at depth limit
            BestResult = Evaluate(Board)
          Else 'NOT (MAXCHAINLENGTH...
            BestResult = -Infinity + Depth               'depth added to find shortest winning line
            For PcMvNum = .MovesListFrom To .MovesListTo
                CurrMoveChain = MoveList(PcMvNum)
                TempBoard = Board
                MakeMove TempBoard, CurrMoveChain, True, False
                If .MovesListTo = 1 Then             'has no choice in depth 0
                    Result = Evaluate(Board)
                    Forced = True
                  Else 'NOT .MOVESLISTTO...
                    With TempBoard
                        .MovesListFrom = .MovesListTo + 1     'establish new segment in movelist
                        .SideToMove = BothSides - .SideToMove 'toggle side to move
                        Result = -Search(TempBoard, Depth + 1, MinDepth, -Beta, -Alpha) 'minimax (negamax) recursion
                        .SideToMove = BothSides - .SideToMove 'back to this side to move
                    End With 'TEMPBOARD
                End If
                If Result > BestResult Then
                    BestResult = Result
                    For i = Depth + 1 To DepthLimit - 1  'create principal variation
                        PV(Depth, i) = PV(Depth + 1, i)  'by copying all best moves after this best move
                        If PV(Depth, i) = "" Then
                            Exit For '>---> Next
                        End If
                    Next i
                    PV(Depth, Depth) = CurrMoveChain     'and enter this best move into PV
                    If BestResult >= Beta Then
                        Cutoffs = Cutoffs + 1
                        Exit For                         'this is bad enough - dont want to know if there are any worse moves '>---> Next
                      ElseIf BestResult > Alpha Then 'NOT BESTRESULT...
                        Alpha = BestResult               '-alpha becomes beta in next depth
                    End If
                End If
            Next PcMvNum
        End If
        Search = BestResult
    End With 'BOARD

End Function

Private Sub tmr_Timer()

    lbMsg = ""
    tmr.Enabled = False

End Sub

Private Sub UnHiliteSquare(ByVal Square As Long)

  Dim Color As Long

    Square = Square - 1
    x = (Square Mod 10) * 64 + 1
    y = (Square \ 10) * 64 + 1
    Color = Point(x - 1, y - 1)
    Line (x, y)-Step(61, 0), Color
    Line -Step(0, 61), Color
    Line -Step(-61, 0), Color
    Line -Step(0, -61), Color

End Sub

':) Ulli's VB Code Formatter V2.10.8 (08.03.2002 11:53:53) 145 + 1201 = 1346 Lines
