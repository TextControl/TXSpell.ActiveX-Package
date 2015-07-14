Imports System.Text
Imports System.Windows.Forms

<ComClass(AxTXSpellChecker.ClassId, AxTXSpellChecker.InterfaceId, AxTXSpellChecker.EventsId)> _
Partial Public Class AxTXSpellChecker

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f4f3f801-a977-498b-816b-9962dbd54156"
    Public Const InterfaceId As String = "cb28e832-f57c-4d16-8cbf-ede7458e25fc"
    Public Const EventsId As String = "bf2ad996-dd16-4399-9e17-66de78f8b486"
#End Region

    ' the main spell checker instance
    ' this wrapper is a proxy to that instance
    Dim txSpellChecker As TXTextControl.Proofing.TXSpellChecker

    Public Sub New()
        ' create a new instance of TX Spell .NET
        txSpellChecker = New TXTextControl.Proofing.TXSpellChecker
        txSpellChecker.Create()
    End Sub

    Public Sub Check(ByVal text As String)
        txSpellChecker.Check(text)
    End Sub

    Public Sub CreateSuggestions(ByVal incorrectWord As String, Optional ByVal max As Integer = Nothing)
        If max = Nothing Then
            txSpellChecker.CreateSuggestions(incorrectWord)
        Else
            txSpellChecker.CreateSuggestions(incorrectWord, max)
        End If
    End Sub

    Public Sub OptionsDialog()
        txSpellChecker.OptionsDialog()
    End Sub

    '*****
    ' MisspelledWordPositions
    '
    ' The MisspelledWordPositions property returns an Integer array
    ' that can be directly used by the SpellCheckText event of
    ' TX Text Control ActiveX
    '*****
    Public ReadOnly Property MisspelledWordPositions() As Integer()
        Get
            If txSpellChecker.IncorrectWords.Count = 0 Then
                Return Nothing
            End If

            Dim lWordPositions(((Me.IncorrectWords.Length - 1) * 2) - 1) As Integer
            Dim i As Integer = 0

            For iCounter As Integer = 0 To Me.IncorrectWords.Length - 2
                Dim word As AxTXSpell.IncorrectWord = Me.IncorrectWords(iCounter)

                lWordPositions.SetValue(word.Start + 1, i)
                lWordPositions.SetValue(word.Start + word.Length, i + 1)
                i = i + 2
            Next

            Return lWordPositions
        End Get
    End Property

    Public Sub AddOpenOfficeDictionary(ByVal OpenOfficeDictionary As OpenOfficeDictionary)
        txSpellChecker.Dictionaries.Add(OpenOfficeDictionary.OriginalDictionary)
    End Sub

    Public Sub AddUserDictionary(ByVal UserDictionary As UserDictionary)
        txSpellChecker.Dictionaries.Add(UserDictionary.OriginalDictionary)
    End Sub

    Public Sub RemoveOpenOfficeDictionary(ByVal OpenOfficeDictionary As OpenOfficeDictionary)
        txSpellChecker.Dictionaries.Remove(OpenOfficeDictionary.OriginalDictionary)
    End Sub

    Public Sub RemoveUserDictionary(ByVal UserDictionary As UserDictionary)
        txSpellChecker.Dictionaries.Remove(UserDictionary.OriginalDictionary)
    End Sub

    Public ReadOnly Property GetOpenOfficeDictionaries() As OpenOfficeDictionary()
        Get

            Dim dicts(txSpellChecker.Dictionaries.Count - 1) As OpenOfficeDictionary
            Dim i As Integer

            For Each dictionary As TXTextControl.Proofing.Dictionary In txSpellChecker.Dictionaries

                Dim newOODictionary As New OpenOfficeDictionary

                If TypeOf dictionary Is TXTextControl.Proofing.OpenOfficeDictionary Then

                    newOODictionary.OriginalDictionary = dictionary

                    If dictionary.DictionaryEncoding Is Encoding.ASCII Then
                        newOODictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.ASCII
                    ElseIf dictionary.DictionaryEncoding Is Encoding.BigEndianUnicode Then
                        newOODictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.BigEndianUnicode
                    ElseIf dictionary.DictionaryEncoding Is Encoding.Default Then
                        newOODictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.DefaultValue
                    ElseIf dictionary.DictionaryEncoding Is Encoding.Unicode Then
                        newOODictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.Unicode
                    ElseIf dictionary.DictionaryEncoding Is Encoding.UTF32 Then
                        newOODictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.UTF32
                    ElseIf dictionary.DictionaryEncoding Is Encoding.UTF7 Then
                        newOODictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.UTF7
                    ElseIf dictionary.DictionaryEncoding Is Encoding.UTF8 Then
                        newOODictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.UTF8
                    End If

                    newOODictionary.IsGetSuggestionsEnabled = dictionary.IsGetSuggestionsEnabled
                    newOODictionary.IsSelectedAsDefault = dictionary.IsSelectedAsDefault
                    newOODictionary.IsSpellCheckingEnabled = dictionary.IsSpellCheckingEnabled
                    newOODictionary.Name = dictionary.Name

                    dicts.SetValue(newOODictionary, i)

                    i = i + 1
                End If
            Next

            Return dicts
        End Get
    End Property

    Public ReadOnly Property GetUserDictionaries() As UserDictionary()
        Get

            Dim dicts(txSpellChecker.Dictionaries.Count - 1) As UserDictionary
            Dim i As Integer

            For Each dictionary As TXTextControl.Proofing.Dictionary In txSpellChecker.Dictionaries

                Dim newUserDictionary As New UserDictionary

                If TypeOf dictionary Is TXTextControl.Proofing.UserDictionary Then

                    newUserDictionary.OriginalDictionary = dictionary

                    If dictionary.DictionaryEncoding Is Encoding.ASCII Then
                        newUserDictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.ASCII
                    ElseIf dictionary.DictionaryEncoding Is Encoding.BigEndianUnicode Then
                        newUserDictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.BigEndianUnicode
                    ElseIf dictionary.DictionaryEncoding Is Encoding.Default Then
                        newUserDictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.DefaultValue
                    ElseIf dictionary.DictionaryEncoding Is Encoding.Unicode Then
                        newUserDictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.Unicode
                    ElseIf dictionary.DictionaryEncoding Is Encoding.UTF32 Then
                        newUserDictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.UTF32
                    ElseIf dictionary.DictionaryEncoding Is Encoding.UTF7 Then
                        newUserDictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.UTF7
                    ElseIf dictionary.DictionaryEncoding Is Encoding.UTF8 Then
                        newUserDictionary.DictionaryEncoding = AxTXSpell.DictionaryEncoding.UTF8
                    End If

                    newUserDictionary.IsGetSuggestionsEnabled = dictionary.IsGetSuggestionsEnabled
                    newUserDictionary.IsSelectedAsDefault = dictionary.IsSelectedAsDefault
                    newUserDictionary.IsSpellCheckingEnabled = dictionary.IsSpellCheckingEnabled
                    newUserDictionary.IsEditable = DirectCast(dictionary, TXTextControl.Proofing.UserDictionary).IsEditable
                    newUserDictionary.Name = dictionary.Name

                    dicts.SetValue(newUserDictionary, i)

                    i = i + 1
                End If
            Next

            Return dicts
        End Get
    End Property

    Public Property IgnoreCase() As IgnoreCaseSettings
        Get
            Return txSpellChecker.IgnoreCase
        End Get
        Set(ByVal value As IgnoreCaseSettings)
            txSpellChecker.IgnoreCase = value
        End Set
    End Property

    Public Property IgnoreWord() As IgnoreWordSettings
        Get
            Return txSpellChecker.IgnoreWord
        End Get
        Set(ByVal value As IgnoreWordSettings)
            txSpellChecker.IgnoreWord = value
        End Set
    End Property

    '*****
    ' Suggestions
    '
    ' Returns a String array of all created suggestions
    '*****
    Public ReadOnly Property GetSuggestions() As String()
        Get
            Dim newSuggestions(txSpellChecker.Suggestions.Count) As String
            Dim i As Integer = 0

            For Each Suggestion As TXTextControl.Proofing.Suggestion In txSpellChecker.Suggestions
                newSuggestions.SetValue(Suggestion.Text, i)
                i = i + 1
            Next

            Return newSuggestions
        End Get
    End Property

    Public ReadOnly Property Suggestions() As Integer
        Get
            Return txSpellChecker.Suggestions.Count
        End Get
    End Property

    Public Property Language() As String
        Get
            Return txSpellChecker.Language
        End Get
        Set(ByVal value As String)
            txSpellChecker.Language = value
        End Set
    End Property

    '*****
    ' IncorrectWords
    '
    ' Returns a array of IncorrectWord
    ' First a proxy IncorrectWord is created/cloned for each
    ' TX Spell .NET IncorrectWord
    '*****
    Public ReadOnly Property IncorrectWords() As IncorrectWord()
        Get
            If txSpellChecker.IncorrectWords Is Nothing Or txSpellChecker.IncorrectWords.Count = 0 Then
                Return Nothing
            End If

            Dim iwIncorrectWords(txSpellChecker.IncorrectWords.Count) As IncorrectWord
            Dim i As Integer = 0

            For Each word As TXTextControl.Proofing.IncorrectWord In txSpellChecker.IncorrectWords
                Dim wrapperIncorrectWord As New IncorrectWord
                wrapperIncorrectWord.Text = word.Text
                wrapperIncorrectWord.Start = word.Start
                wrapperIncorrectWord.Length = word.Length
                wrapperIncorrectWord.IsDuplicate = word.IsDuplicate

                iwIncorrectWords.SetValue(wrapperIncorrectWord, i)
                i = i + 1
            Next

            Return iwIncorrectWords
        End Get
    End Property
End Class







