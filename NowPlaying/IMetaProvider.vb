Public Interface IMetaProvider
    Inherits IDisposable

    Function GetState() As ProviderState
    Function SupportsMeta(field As MetaField) As Boolean
    Function GenerateMeta(field As MetaField) As String

End Interface