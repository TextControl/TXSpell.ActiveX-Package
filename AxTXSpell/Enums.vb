Public Enum IgnoreCaseSettings
    AllLower = 1
    AllUpper = 2
    Always = 4
    Never = 8
    WordBegin = 32
    WordBeginUpper = 64
End Enum

Public Enum IgnoreWordSettings
    InUpperCase = 1
    Never = 2
    WithNumbers = 8
    Duplicated = 16
    IsEmail = 32
    IsURL = 64
End Enum

Public Enum DictionaryEncoding
    ASCII = 1
    BigEndianUnicode = 2
    DefaultValue = 4
    Unicode = 8
    UTF32 = 16
    UTF7 = 32
    UTF8 = 64
End Enum