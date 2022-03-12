Imports System
Imports System.Collections.Generic
Imports System.Management

''' <summary>
''' Contains functions to simplify queries against the WMI database.
''' </summary>
''' <remarks></remarks>
Public Class WMI

    Private Shared cimTypeTable As Dictionary(Of CimType, Type) = InitializeTypeTable()

    Private Shared Function InitializeTypeTable() As Dictionary(Of CimType, Type)
        Dim table As New Dictionary(Of CimType, Type)

        table.Add(CimType.Boolean, GetType(Boolean))
        table.Add(CimType.Char16, GetType(String))
        table.Add(CimType.DateTime, GetType(DateTime))
        table.Add(CimType.Object, GetType(Object))
        table.Add(CimType.Real32, GetType(Decimal))
        table.Add(CimType.Real64, GetType(Decimal))
        table.Add(CimType.Reference, GetType(Object))
        table.Add(CimType.SInt8, GetType(SByte))
        table.Add(CimType.SInt16, GetType(Short))
        table.Add(CimType.SInt32, GetType(Integer))
        table.Add(CimType.SInt64, GetType(Long))
        table.Add(CimType.String, GetType(String))
        table.Add(CimType.UInt8, GetType(Byte))
        table.Add(CimType.UInt16, GetType(UShort))
        table.Add(CimType.UInt32, GetType(UInteger))
        table.Add(CimType.UInt64, GetType(ULong))

        Return table
    End Function

    Public Function BuildNamespaceList() As List(Of String)
        Return Me.BuildNamespaceList("root")
    End Function

    Public Function BuildNamespaceList(selectedNamespace As String) As List(Of String)
        Dim nsList As New List(Of String)

        Using nsClass As New ManagementClass(selectedNamespace, "__namespace", Nothing)
            Using nsCollection As ManagementObjectCollection = nsClass.GetInstances()
                For Each ns As ManagementObject In nsCollection
                    nsList.Add(ns("Name"))
                Next
            End Using
        End Using

        Return nsList
    End Function

    Public Function BuildClassList(selectedNamespace As String) As List(Of String)
        Dim clList As New List(Of String)

        Using clSearch As New ManagementObjectSearcher(selectedNamespace, "select * from meta_class")
            Using clCollection As ManagementObjectCollection = clSearch.Get()
                For Each clClass As ManagementClass In clCollection
                    For Each clQualifier As QualifierData In clClass.Qualifiers
                        If clQualifier.Name.Equals("dynamic") OrElse clQualifier.Name.Equals("static") Then
                            clList.Add(clClass("__CLASS").ToString())
                        End If
                    Next
                Next
            End Using
        End Using

        Return clList
    End Function

    Public Function BuildPropertyList(selectedNamespace As String, selectedClass As String) As PropertyDataCollection
        Dim prOptions As New ObjectGetOptions(Nothing, TimeSpan.MaxValue, True)
        Dim prList As PropertyDataCollection

        Using prClass As New ManagementClass(selectedNamespace, selectedClass, prOptions)
            prList = prClass.Properties
        End Using

        Return prList
    End Function

    Public Function GetPropertyType(selectedProperty As PropertyData) As Type
        Dim prType As Type = WMI.cimTypeTable(selectedProperty.Type)

        If selectedProperty.IsArray Then
            Return prType.MakeArrayType()
        Else
            Return prType
        End If
    End Function

    Public Function GetPropertyValue(selectedProperty As PropertyData) As Object


        Select Case selectedProperty.Type

            Case CimType.DateTime
                Return ManagementDateTimeConverter.ToDateTime(selectedProperty.Value.ToString())

            Case Else
                Return Convert.ChangeType(selectedProperty.Value, GetPropertyType(selectedProperty))

        End Select
    End Function

    Public Function GetInstances(selectedClass As String) As ManagementObjectCollection
        Dim query As String = "select * from " + selectedClass

        Using searcher As New ManagementObjectSearcher(query)
            Return searcher.Get()
        End Using
    End Function

End Class