Attribute VB_Name = "Declaration"
Global CurString As String

Global typeDevice As String
Global AliasName As String
Global Result As String
Global Started As Boolean
Global PlayMode As String

Public Type MP3Info
    Bitrate As Integer
    Frequency As Long
    Mode As String
    Emphasis As String
    'ModeExtension As String
    MpegVersion As Integer
    MpegLayer As Integer
    Padding As String
    CRC As String
    Duration As Long
    CopyRight As String
    Original As String
    PrivateBit As String
    HasTag As Boolean
    Tag As String
    Songname As String
    Artist As String
    Album As String
    Year As String
    Comment As String
    Genre As Integer
    Track As String
    VBR As Boolean
    Frames As Integer
End Type

Global GetMP3Info As MP3Info

