<ComClass(IncorrectWord.ClassId, IncorrectWord.InterfaceId, IncorrectWord.EventsId)> _
Public Class IncorrectWord

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "b10a7522-bf8b-4fe0-bf00-4e3d00b55d2a"
    Public Const InterfaceId As String = "c5951ce8-bf9e-4156-8053-db7dc8045456"
    Public Const EventsId As String = "65793074-0092-429f-9951-3702166cdfc9"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.

    Private sText As String
    Private iStart As Integer
    Private iLength As Integer
    Private bIsDuplicate As Boolean

    Public Property Text() As String
        Get
            Return sText
        End Get
        Set(ByVal value As String)
            sText = value
        End Set
    End Property

    Public Property IsDuplicate() As Boolean
        Get
            Return bIsDuplicate
        End Get
        Set(ByVal value As Boolean)
            bIsDuplicate = value
        End Set
    End Property

    Public Property Start() As Integer
        Get
            Return iStart
        End Get
        Set(ByVal value As Integer)
            iStart = value
        End Set
    End Property

    Public Property Length() As Integer
        Get
            Return iLength
        End Get
        Set(ByVal value As Integer)
            iLength = value
        End Set
    End Property



End Class



