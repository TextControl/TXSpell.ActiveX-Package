VERSION 5.00
Object = "{0a8ef900-46e5-11e3-a545-0013d350667c}#2.6#0"; "tx4ole20.ocx"
Begin VB.Form Form1 
   Caption         =   "TX Spell .NET ActiveX Package Sample"
   ClientHeight    =   8640
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Spelling..."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   8040
      Width           =   1815
   End
   Begin Tx4oleLib.TXTextControl TXTextControl1 
      Align           =   1  'Align Top
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _Version        =   131075
      _ExtentX        =   18230
      _ExtentY        =   13996
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
      ViewMode        =   2
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
   End
   Begin VB.Menu mnu_Spell 
      Caption         =   "Spell"
      Visible         =   0   'False
      Begin VB.Menu mnuWord1 
         Caption         =   "-xcv"
      End
      Begin VB.Menu mnuWord2 
         Caption         =   "-xcv"
      End
      Begin VB.Menu mnuWord3 
         Caption         =   "-xcv"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDialog 
         Caption         =   "Spelling Dialog..."
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options..."
      End
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

Private Sub Command2_Click()
    Form2.TXSpellChecker = TXSpell1
    Form2.TextControl = TXTextControl1
    
    Form2.Show
End Sub

Private Sub mnuDialog_Click()
    Form2.TXSpellChecker = TXSpell1
    Form2.TextControl = TXTextControl1
    
    Form2.Show
End Sub

Private Sub mnuOptions_Click()
    TXSpell1.OptionsDialog
End Sub

Private Sub mnuWord1_Click()
    TXTextControl1.MisspelledWordDelete TXTextControl1.MisspelledWordAtInputPos, mnuWord1.Caption
End Sub

Private Sub mnuWord2_Click()
   TXTextControl1.MisspelledWordDelete TXTextControl1.MisspelledWordAtInputPos, mnuWord2.Caption
End Sub

Private Sub mnuWord3_Click()
   TXTextControl1.MisspelledWordDelete TXTextControl1.MisspelledWordAtInputPos, mnuWord3.Caption
End Sub

Private Sub TXTextControl1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
        
    If Button = 2 Then
        TXTextControl1.SelStart = TXTextControl1.InputPosFromPoint(X, Y)
    End If
        
    If TXTextControl1.MisspelledWordAtInputPos = 0 Or Button <> 2 Then
        Exit Sub
    End If
        
    TXTextControl1.SelStart = TXTextControl1.MisspelledWordStart(TXTextControl1.MisspelledWordAtInputPos)
    TXTextControl1.SelLength = TXTextControl1.MisspelledWordLength(TXTextControl1.MisspelledWordAtInputPos)
    
    TXSpell1.CreateSuggestions (TXTextControl1.SelText)
    
    suggestionsArray = TXSpell1.GetSuggestions
    
    For i = 0 To 3
        Select Case i
            Case 0
            mnuWord1.Caption = suggestionsArray(0)
            
            If mnuWord1.Caption = "" Then
                mnuWord1.Caption = "No suggestions found."
            End If
                      
            Case 1
            mnuWord2.Caption = ""
            mnuWord2.Caption = suggestionsArray(1)
            
            If mnuWord2.Caption = "" Then
                mnuWord2.Visible = False
            Else
                mnuWord2.Visible = True
            End If
            
            
            Case 2
            mnuWord3.Caption = ""
            mnuWord3.Caption = suggestionsArray(2)
            
            If mnuWord3.Caption = "" Then
                mnuWord3.Visible = False
            Else
                mnuWord3.Visible = True
            End If
        End Select
    Next
        
    PopupMenu mnu_Spell
    
End Sub

Private Sub TXTextControl1_SpellCheckText(ByVal Text As String, MisspelledWordPositions As Variant)
    TXSpell1.Check (Text)
    MisspelledWordPositions = TXSpell1.MisspelledWordPositions
End Sub
