Imports System.Text

<ComClass(OpenOfficeDictionary.ClassId, OpenOfficeDictionary.InterfaceId, OpenOfficeDictionary.EventsId)> _
Public Class OpenOfficeDictionary

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ed617dd6-0dea-413d-b00f-de103fe9351f"
    Public Const InterfaceId As String = "785d161a-92f5-4e0c-afab-6f71f61b8dcb"
    Public Const EventsId As String = "29fc8241-4be6-4aaa-afca-7c4de30cebc0"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.

    Private sName As String
    Private enDictionaryEncoding As DictionaryEncoding
    Private bIsGetSuggestionsEnabled As Boolean
    Private bIsSelectedAsDefault As Boolean
    Private bIsSpellCheckingEnabled As Boolean
    Private sDictionaryPath As String = String.Empty
    Private sLanguage As String = String.Empty
    Private dOriginalDictionary As TXTextControl.Proofing.OpenOfficeDictionary

    Public Sub New()
    End Sub

    Public Sub New(ByVal userDictionaryPath As String)
        DictionaryPath = userDictionaryPath
        OriginalDictionary = New TXTextControl.Proofing.OpenOfficeDictionary(userDictionaryPath)
    End Sub

    Friend Property OriginalDictionary() As TXTextControl.Proofing.OpenOfficeDictionary
        Get
            Return dOriginalDictionary
        End Get
        Set(ByVal value As TXTextControl.Proofing.OpenOfficeDictionary)
            dOriginalDictionary = value
        End Set
    End Property

    Public Property DictionaryPath() As String
        Get
            Return sDictionaryPath
        End Get
        Set(ByVal value As String)
            sDictionaryPath = value
            OriginalDictionary = New TXTextControl.Proofing.OpenOfficeDictionary(sDictionaryPath)
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
    Public Property Language() As String
        Get
            Return sLanguage
        End Get
        Set(ByVal value As String)
            sLanguage = value
            OriginalDictionary.Language = New CultureInfo(sLanguage)
        End Set
    End Property
End Class


