VERSION 5.00
Object = "{67d06a00-8469-11e5-a5c5-0013d350667c}#2.9#0"; "tx4ole23.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spelling"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOptions 
      Caption         =   "Optio&ns..."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton btnChangeAll 
      Caption         =   "C&hange All"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton btnChange 
      Caption         =   "Chan&ge"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ListBox listSuggestions 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   5295
   End
   Begin VB.CommandButton btnIgnoreAll 
      Caption         =   "Ign&ore All"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton btnIgnoreOnce 
      Caption         =   "Ignore On&ce"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin Tx4oleLib.TXTextControl txcNotInDictionaries 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
      _Version        =   131075
      _ExtentX        =   9340
      _ExtentY        =   2355
      _StockProps     =   73
      BackColor       =   16777215
      Language        =   1
      BorderStyle     =   1
      BackStyle       =   1
      ControlChars    =   0   'False
      EditMode        =   1
      HideSelection   =   -1  'True
      InsertionMode   =   -1  'True
      MousePointer    =   0
      ZoomFactor      =   100
      ViewMode        =   3
      ClipChildren    =   0   'False
      ClipSiblings    =   -1  'True
      SizeMode        =   0
      TabKey          =   -1  'True
      FormatSelection =   -1  'True
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
      Text            =   ""
      WordWrapMode    =   1
      AllowUndo       =   -1  'True
      TextFrameMarkerLines=   -1  'True
      FieldLinkTargetMarkers=   0   'False
      PageOrientation =   0
      PageViewStyle   =   1
      FontSettings    =   0
      AllowDrag       =   0   'False
      AllowDrop       =   0   'False
      SelectionViewMode=   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Suggestions:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Not in &Dictionaries:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private txSpell As AxTXSpellChecker
Private tx As Tx4oleLib.TXTextControl
Dim misspelledWord As String

Public Property Let TXSpellChecker(ByVal vNewValue As AxTXSpellChecker)
    Set txSpell = vNewValue
End Property

Public Property Let TextControl(ByVal vNewValue As Tx4oleLib.TXTextControl)
    Set tx = vNewValue
End Property

Private Sub Command6_Click()
    Form2.Hide
End Sub

Private Sub Command7_Click()
    txSpell.OptionsDialog
End Sub

Private Sub CloseDialog()
    Me.Hide
End Sub

Private Sub NextMisspelledWord()
    If tx.MisspelledWords = 0 Then
        MsgBox "Spell checking complete.", vbOKOnly, "TX Spell .NET ActiveX Package"
        CloseDialog
    Else
        tx.SelStart = tx.MisspelledWordStart(1)
        tx.SelLength = tx.MisspelledWordLength(1)
        
        misspelledWord = tx.SelText
        
        txSpell.Check (misspelledWord)
        txSpell.CreateSuggestions misspelledWord, 10
        
        Suggestions = txSpell.GetSuggestions
        
        listSuggestions.Clear
        
        For i = 0 To txSpell.Suggestions
            listSuggestions.AddItem (Suggestions(i))
        Next
        
        If listSuggestions.ListCount > 1 Then
            btnChange.Enabled = True
            btnChangeAll.Enabled = True
            listSuggestions.ListIndex = 0
        Else
            btnChange.Enabled = False
            btnChangeAll.Enabled = False
            listSuggestions.ListIndex = -1
        End If
        
        tx.SelStart = tx.GetCharFromLine(tx.GetLineFromChar(tx.SelStart))
        tx.SelLength = tx.MisspelledWordStart(1) + 20
        
        txcNotInDictionaries.Text = tx.SelText
        
        txcNotInDictionaries.SelStart = tx.MisspelledWordStart(1) - tx.SelStart
        txcNotInDictionaries.SelLength = tx.MisspelledWordLength(1)
        txcNotInDictionaries.ForeColor = vbRed
        FieldText = txcNotInDictionaries.SelText
        
        txcNotInDictionaries.SelText = ""
        
        txcNotInDictionaries.SelLength = 0
        txcNotInDictionaries.FieldInsert (FieldText)
        txcNotInDictionaries.FieldChangeable = False
        
        tx.SelStart = tx.MisspelledWordStart(1)
        tx.SelLength = tx.MisspelledWordLength(1)
    End If
End Sub


Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnChange_Click()
    tx.MisspelledWordDelete 1, listSuggestions.Text
    NextMisspelledWord
End Sub

Private Sub btnChangeAll_Click()
    ChangeAll (listSuggestions.Text)
End Sub

Private Sub btnIgnoreAll_Click()
    ChangeAll (misspelledWord)
End Sub

Private Sub ChangeAll(correctedWord As String)
    Dim iCounter As Integer
    iCounter = 1

    For i = 1 To tx.MisspelledWords
        tx.SelStart = tx.MisspelledWordStart(iCounter)
        tx.SelLength = tx.MisspelledWordLength(iCounter)
        
        If tx.SelText = misspelledWord Then
            tx.MisspelledWordDelete iCounter, misspelledWord
        Else
            iCounter = iCounter + 1
        End If
    Next
    
    NextMisspelledWord
End Sub

Private Sub btnIgnoreOnce_Click()
    tx.MisspelledWordDelete 1, misspelledWord
    NextMisspelledWord
End Sub

Private Sub Form_Load()
    tx.HideSelection = False
    NextMisspelledWord
End Sub

Private Sub txcNotInDictionaries_PosChange()
    If txcNotInDictionaries.FieldAtInputPos <> 0 Then
        txcNotInDictionaries.EditMode = 0
    Else
        txcNotInDictionaries.EditMode = 1
    End If
End Sub
