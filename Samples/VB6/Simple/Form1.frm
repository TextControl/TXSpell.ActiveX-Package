VERSION 5.00
Object = "{0A8EF900-46E5-11E3-A545-0013D350667C}#2.6#0"; "tx4ole20.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin Tx4oleLib.TXTextControl TXTextControl1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _Version        =   131078
      _ExtentX        =   16960
      _ExtentY        =   9551
      _StockProps     =   73
      BackColor       =   16777215
      Language        =   1
      BorderStyle     =   1
      BackStyle       =   1
      ControlChars    =   0   'False
      EditMode        =   0
      HideSelection   =   -1  'True
      InsertionMode   =   -1  'True
      MousePointer    =   0
      ZoomFactor      =   100
      ViewMode        =   3
      ClipChildren    =   0   'False
      ClipSiblings    =   -1  'True
      SizeMode        =   0
      TabKey          =   -1  'True
      FormatSelection =   0   'False
      VTSpellDictionary=   "C:\PROGRA~2\TEXTCO~1\TXTEXT~1.0AC\Bin\AMERICAN.VTD"
      ScrollBars      =   3
      PageWidth       =   12240
      PageHeight      =   15840
      PageMarginL     =   1440
      PageMarginT     =   1440
      PageMarginR     =   1440
      PageMarginB     =   1440
      PrintZoom       =   100
      PrintOffset     =   0   'False
      PrintColors     =   -1  'True
      FontName        =   "Arial"
      FontSize        =   12
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Baseline        =   0
      TextBkColor     =   16777215
      Alignment       =   0
      LineSpacing     =   100
      LineSpacingT    =   0
      FrameStyle      =   32
      FrameDistance   =   0
      FrameLineWidth  =   20
      IndentL         =   0
      IndentR         =   0
      IndentFL        =   0
      IndentT         =   0
      IndentB         =   0
      Text            =   "TXTextControl1"
      WordWrapMode    =   1
      AllowUndo       =   -1  'True
      TextFrameMarkerLines=   -1  'True
      FieldLinkTargetMarkers=   0   'False
      PageOrientation =   0
      PageViewStyle   =   1
      FontSettings    =   0
      AllowDrag       =   0   'False
      AllowDrop       =   0   'False
      EnableSpellChecking=   -1  'True
      SelectionViewMode=   1
      SectionRestartPageNumbering=   0
      PermanentControlChars=   16
      RightToLeft     =   0   'False
      TextDirection   =   2
      Locale          =   1033
      Justification   =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXSpell1 As AxTXSpell.AxTXSpellChecker

Private Sub Form_Load()
    Set TXSpell1 = New AxTXSpellChecker
End Sub

Private Sub TXTextControl1_SpellCheckText(ByVal Text As String, MisspelledWordPositions As Variant)
    TXSpell1.Check (Text)
    MisspelledWordPositions = TXSpell1.MisspelledWordPositions
End Sub
