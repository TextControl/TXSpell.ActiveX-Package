Imports System.Text

<ComClass(UserDictionary.ClassId, UserDictionary.InterfaceId, UserDictionary.EventsId)> _
Public Class UserDictionary

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "b90e1d28-82d0-4536-8cee-71c7665f22fa"
    Public Const InterfaceId As String = "bf0d7908-c081-4b40-9a28-8fec082d58b4"
    Public Const EventsId As String = "0595c562-82ba-4a71-8e28-d1e69e199187"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.

    Private bIsEditable As Boolean
    Private sName As String
    Private enDictionaryEncoding As DictionaryEncoding
    Private bIsGetSuggestionsEnabled As Boolean
    Private bIsSelectedAsDefault As Boolean
    Private bIsSpellCheckingEnabled As Boolean
    Private sDictionaryPath As String = String.Empty
    Private dOriginalDictionary As TXTextControl.Proofing.UserDictionary

    Public Sub New()
        OriginalDictionary = New TXTextControl.Proofing.UserDictionary()
    End Sub

    Public Sub New(ByVal userDictionaryPath As String)
        DictionaryPath = userDictionaryPath
        OriginalDictionary = New TXTextControl.Proofing.UserDictionary(userDictionaryPath)
    End Sub

    Public Property OriginalDictionary() As TXTextControl.Proofing.Dictionary
        Get
            Return dOriginalDictionary
        End Get
        Set(ByVal value As TXTextControl.Proofing.Dictionary)
            dOriginalDictionary = value
        End Set
    End Property

    Public Property DictionaryPath() As String
        Get
            Return sDictionaryPath
        End Get
        Set(ByVal value As String)
            sDictionaryPath = value
            OriginalDictionary = New TXTextControl.Proofing.UserDictionary(sDictionaryPath)
        End Set
    End Property

    Public Property DictionaryEncoding() As DictionaryEncoding
        Get
            Return enDictionaryEncoding
        End Get
        Set(ByVal value As DictionaryEncoding)
            enDictionaryEncoding = value

            If value = AxTXSpell.DictionaryEncoding.ASCII Then
                OriginalDictionary.DictionaryEncoding = Encoding.ASCII
            ElseIf value = AxTXSpell.DictionaryEncoding.BigEndianUnicode Then
                OriginalDictionary.DictionaryEncoding = Encoding.BigEndianUnicode
            ElseIf value = AxTXSpell.DictionaryEncoding.DefaultValue Then
                OriginalDictionary.DictionaryEncoding = Encoding.Default
            ElseIf value = AxTXSpell.DictionaryEncoding.Unicode Then
                OriginalDictionary.DictionaryEncoding = Encoding.Unicode
            ElseIf value = AxTXSpell.DictionaryEncoding.UTF32 Then
                OriginalDictionary.DictionaryEncoding = Encoding.UTF32
            ElseIf value = AxTXSpell.DictionaryEncoding.UTF7 Then
                OriginalDictionary.DictionaryEncoding = Encoding.UTF7
            ElseIf value = AxTXSpell.DictionaryEncoding.UTF8 Then
                OriginalDictionary.DictionaryEncoding = Encoding.UTF8
            End If

        End Set
    End Property

    Public Property Name() As String
        Get
            Return sName
        End Get
        Set(ByVal value As String)
            sName = value
            OriginalDictionary.Name = value
        End Set
    End Property

    Public Property IsGetSuggestionsEnabled() As Boolean
        Get
            Return bIsGetSuggestionsEnabled
        End Get
        Set(ByVal value As Boolean)
            bIsGetSuggestionsEnabled = value
            OriginalDictionary.IsGetSuggestionsEnabled = value
        End Set
    End Property

    Public Property IsSelectedAsDefault() As Boolean
        Get
            Return bIsSelectedAsDefault
        End Get
        Friend Set(ByVal value As Boolean)
            bIsSelectedAsDefault = value
        End Set
    End Property

    Public Property IsSpellCheckingEnabled() As Boolean
        Get
            Return bIsSpellCheckingEnabled
        End Get
        Set(ByVal value As Boolean)
            bIsSpellCheckingEnabled = value
            OriginalDictionary.IsSpellCheckingEnabled = value
        End Set
    End Property

    Public Property IsEditable() As Boolean
        Get
            Return bIsEditable
        End Get
        Set(ByVal value As Boolean)
            bIsEditable = value

            Dim dic As TXTextControl.Proofing.UserDictionary = OriginalDictionary
            dic.IsEditable = value
        End Set
    End Property

    Public Function AddWord(ByVal word As String) As Boolean
        Dim dic As TXTextControl.Proofing.UserDictionary = OriginalDictionary
        Return dic.AddWord(word)
    End Function

    Public Function RemoveAllWords() As Boolean
        Dim dic As TXTextControl.Proofing.UserDictionary = OriginalDictionary
        Return dic.RemoveAllWords
    End Function

    Public Function RemoveWord(ByVal word As String) As Boolean
        Dim dic As TXTextControl.Proofing.UserDictionary = OriginalDictionary
        Return dic.RemoveWord(word)
    End Function

    Public Sub Save(ByVal userDictionaryPath As String)
        Dim dic As TXTextControl.Proofing.UserDictionary = OriginalDictionary
        dic.Save(userDictionaryPath)
    End Sub

    Public Function ToArray() As String()
        Dim dic As TXTextControl.Proofing.UserDictionary = OriginalDictionary
        Return dic.ToArray()
    End Function

End Class


