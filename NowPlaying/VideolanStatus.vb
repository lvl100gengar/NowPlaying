Imports System.Xml.Serialization

<XmlRoot("root")> _
Public Class VideolanStatus

    <XmlElement("volume")> _
    Public Property Volume As Integer
    <XmlElement("length")> _
    Public Property Length As Integer
    <XmlElement("time")> _
    Public Property Time As Integer
    <XmlElement("state")> _
    Public Property State As String
    <XmlElement("position")> _
    Public Property Position As Integer
    <XmlElement("fullscreen")> _
    Public Property FullScreen As String
    <XmlElement("random")> _
    Public Property Random As Integer
    <XmlElement("loop")> _
    Public Property [Loop] As Integer
    <XmlElement("repeat")> _
    Public Property Repeat As String
    <XmlElement("information")> _
    Public Property Information As VideolanMediaInfo
    <XmlElement("stats")> _
    Public Property Stats As VideolanStats

End Class

Public Class VideolanMediaInfo

    <XmlElement("category")> _
    Public Property Streams As VideolanStreamInfo()
    <XmlElement("meta-information")> _
    Public Property MetaInfo As VideolanMetaInfo

    Public ReadOnly Property Streams(ByVal streamType As String) As VideolanStreamInfo
        Get
            'Make sure the streams property was initialized.
            If Me.Streams Is Nothing Then
                Return Nothing
            End If

            'Look for a stream with the specified type.
            For Each availableStream As VideolanStreamInfo In Me.Streams
                If availableStream.Attributes IsNot Nothing AndAlso availableStream.Attributes("Type").Equals(streamType) Then
                    Return availableStream
                End If
            Next

            'Return null if no stream of the specified type was found.
            Return Nothing
        End Get
    End Property

End Class

Public Class VideolanStreamInfo

    <XmlAttribute("name")> _
    Public Property Name As String
    <XmlElement("info")> _
    Public Property Attributes As VideolanStreamAttribute()

    <XmlIgnore()> _
    Public ReadOnly Property Attributes(ByVal attributeName As String) As String
        Get
            'Make sure the attributes array was initialized.
            If Me.Attributes Is Nothing Then
                Return String.Empty
            End If

            'Find the attribute with the given name.
            For Each availableAttribute As VideolanStreamAttribute In Me.Attributes
                'Skip attributes that are null or have no name.
                If availableAttribute Is Nothing OrElse availableAttribute.Name Is Nothing Then
                    Continue For
                End If

                'See if the name of this attribute matches the input name.
                If availableAttribute.Name.Equals(attributeName) Then
                    Return If(availableAttribute.Value, String.Empty)
                End If
            Next

            'Return an empty String if the specified attribute was not found.
            Return String.Empty
        End Get
    End Property

End Class

Public Class VideolanStreamAttribute

    <XmlAttribute("name")> _
    Public Property Name As String
    <XmlText()> _
    Public Property Value As String

End Class

Public Class VideolanMetaInfo

    <XmlElement("title")> _
    Public Property Title As String
    <XmlElement("artist")> _
    Public Property Artist As String
    <XmlElement("genre")> _
    Public Property Genre As String
    <XmlElement("copyright")> _
    Public Property Copyright As String
    <XmlElement("album")> _
    Public Property Album As String
    <XmlElement("track")> _
    Public Property Track As String
    <XmlElement("description")> _
    Public Property Description As String
    <XmlElement("rating")> _
    Public Property Rating As String
    <XmlElement("date")> _
    Public Property [Date] As String
    <XmlElement("url")> _
    Public Property URL As String
    <XmlElement("language")> _
    Public Property Language As String
    <XmlElement("now_playing")> _
    Public Property NowPlaying As String
    <XmlElement("publisher")> _
    Public Property Publisher As String
    <XmlElement("encoded_by")> _
    Public Property Encoder As String
    <XmlElement("art_url")> _
    Public Property ArtURL As String
    <XmlElement("track_id")> _
    Public Property TrackID As String

End Class

Public Class VideolanStats

    <XmlElement("readbytes")> _
    Public Property ReadBytes As String
    <XmlElement("inputbitrate")> _
    Public Property InputBitrate As String
    <XmlElement("demuxreadbytes")> _
    Public Property DemuxReadBytes As String
    <XmlElement("demuxbitrate")> _
    Public Property DemuxBitrate As String
    <XmlElement("decodedvideo")> _
    Public Property DecodedVideo As String
    <XmlElement("displayedpictures")> _
    Public Property DisplayedPictures As String
    <XmlElement("lostpictures")> _
    Public Property LostPictures As String
    <XmlElement("decodedaudio")> _
    Public Property DecodedAudio As String
    <XmlElement("playedabuffers")> _
    Public Property playedabuffers As String
    <XmlElement("lostabuffers")> _
    Public Property LostaBuffers As String
    <XmlElement("sentpackets")> _
    Public Property SentPackets As String
    <XmlElement("sentbytes")> _
    Public Property SentBytes As String
    <XmlElement("sendbitrate")> _
    Public Property SendBitrate As String

End Class

<XmlRoot("node")> _
Public Class VideolanNode

    <XmlAttribute("id")> _
    Public Property ID As Integer
    <XmlAttribute("name")> _
    Public Property Name As String
    <XmlAttribute("ro")> _
    Public Property RO As String
    <XmlElement("node")> _
    Public Property Nodes As VideolanNode()
    <XmlElement("leaf")> _
    Public Property Items As VideolanItem()

    Public Function FindCurrent(ByRef currentPlaylist As VideolanNode, ByRef currentItem As VideolanItem) As Boolean
        'Make sure the Nodes property was initialized properly.
        If Me.Nodes IsNot Nothing Then
            'Check my children to see if they have the current playlist/item.
            For Each childNode As VideolanNode In Me.Nodes
                If childNode.FindCurrent(currentPlaylist, currentItem) Then
                    Return True
                End If
            Next
        End If

        'Make sure the Items property was initialized properly.
        If Me.Items IsNot Nothing Then
            For Each childItem As VideolanItem In Me.Items
                If childItem.IsCurrentItem Then
                    currentPlaylist = Me
                    currentItem = childItem
                    Return True
                End If
            Next
        End If

        Return False
    End Function

End Class

Public Class VideolanItem
    Inherits VideolanMetaInfo

    <XmlAttribute("id")> _
    Public Property ID As Integer
    <XmlAttribute("current")> _
    Public Property Current As String
    <XmlAttribute("uri")> _
    Public Property URI As String
    <XmlAttribute("name")> _
    Public Property Name As String
    <XmlAttribute("ro")> _
    Public Property RO As String
    <XmlAttribute("duration")> _
    Public Property Duration As String

    Public ReadOnly Property IsCurrentItem() As Boolean
        Get
            Return Me.Current IsNot Nothing AndAlso Me.Current.Equals("current")
        End Get
    End Property

End Class