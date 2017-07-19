﻿
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Imports System
Imports System.Data.Objects
Imports System.Data.Objects.DataClasses
Imports System.Data.EntityClient
Imports System.ComponentModel
Imports System.Xml.Serialization
Imports System.Runtime.Serialization

<Assembly: EdmSchemaAttribute("4ba7a837-800c-42ca-8442-89ea0c25e1c3")>
#Region "EDM Relationship Metadata"
<Assembly: EdmRelationshipAttribute("TPHModel", "TPHPrimaryDataTPHSecondaryData", "TPHPrimaryData", System.Data.Metadata.Edm.RelationshipMultiplicity.One, GetType(TPHPrimaryData), "TPHSecondaryData", System.Data.Metadata.Edm.RelationshipMultiplicity.Many, GetType(TPHSecondaryData), True)>

#End Region

#Region "Contexts"

''' <summary>
''' No Metadata Documentation available.
''' </summary>
Public Partial Class TPHModelContainer
    Inherits ObjectContext

    #Region "Constructors"

    ''' <summary>
    ''' Initializes a new TPHModelContainer object using the connection string found in the 'TPHModelContainer' section of the application configuration file.
    ''' </summary>
    Public Sub New()
        MyBase.New("name=TPHModelContainer", "TPHModelContainer")
    MyBase.ContextOptions.LazyLoadingEnabled = true
        OnContextCreated()
    End Sub

    ''' <summary>
    ''' Initialize a new TPHModelContainer object.
    ''' </summary>
    Public Sub New(ByVal connectionString As String)
        MyBase.New(connectionString, "TPHModelContainer")
    MyBase.ContextOptions.LazyLoadingEnabled = true
        OnContextCreated()
    End Sub

    ''' <summary>
    ''' Initialize a new TPHModelContainer object.
    ''' </summary>
    Public Sub New(ByVal connection As EntityConnection)
        MyBase.New(connection, "TPHModelContainer")
    MyBase.ContextOptions.LazyLoadingEnabled = true
        OnContextCreated()
    End Sub

    #End Region

    #Region "Partial Methods"

    Partial Private Sub OnContextCreated()
    End Sub

    #End Region

    #Region "ObjectSet Properties"

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    Public ReadOnly Property TPHDataItems() As ObjectSet(Of TPHDataItem)
        Get
            If (_TPHDataItems Is Nothing) Then
                _TPHDataItems = MyBase.CreateObjectSet(Of TPHDataItem)("TPHDataItems")
            End If
            Return _TPHDataItems
        End Get
    End Property

    Private _TPHDataItems As ObjectSet(Of TPHDataItem)

    #End Region
    #Region "AddTo Methods"

    ''' <summary>
    ''' Deprecated Method for adding a new object to the TPHDataItems EntitySet. Consider using the .Add method of the associated ObjectSet(Of T) property instead.
    ''' </summary>
    Public Sub AddToTPHDataItems(ByVal tPHDataItem As TPHDataItem)
        MyBase.AddObject("TPHDataItems", tPHDataItem)
    End Sub

    #End Region
End Class

#End Region
#Region "Entities"

''' <summary>
''' No Metadata Documentation available.
''' </summary>
<EdmEntityTypeAttribute(NamespaceName:="TPHModel", Name:="TPHDataItem")>
<Serializable()>
<DataContractAttribute(IsReference:=True)>
<KnownTypeAttribute(GetType(TPHPrimaryData))>
<KnownTypeAttribute(GetType(TPHSecondaryData))>
Public Partial Class TPHDataItem
    Inherits EntityObject
    #Region "Factory Method"

    ''' <summary>
    ''' Create a new TPHDataItem object.
    ''' </summary>
    ''' <param name="index">Initial value of the Index property.</param>
    ''' <param name="name">Initial value of the Name property.</param>
    ''' <param name="value">Initial value of the Value property.</param>
    Public Shared Function CreateTPHDataItem(index As Global.System.Int32, name As Global.System.String, value As Global.System.String) As TPHDataItem
        Dim tPHDataItem as TPHDataItem = New TPHDataItem
        tPHDataItem.Index = index
        tPHDataItem.Name = name
        tPHDataItem.Value = value
        Return tPHDataItem
    End Function

    #End Region
    #Region "Primitive Properties"

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <EdmScalarPropertyAttribute(EntityKeyProperty:=true, IsNullable:=false)>
    <DataMemberAttribute()>
    Public Property Index() As Global.System.Int32
        Get
            Return _Index
        End Get
        Set
            If (_Index <> Value) Then
                OnIndexChanging(value)
                ReportPropertyChanging("Index")
                _Index = StructuralObject.SetValidValue(value)
                ReportPropertyChanged("Index")
                OnIndexChanged()
            End If
        End Set
    End Property

    Private _Index As Global.System.Int32
    Private Partial Sub OnIndexChanging(value As Global.System.Int32)
    End Sub

    Private Partial Sub OnIndexChanged()
    End Sub

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <EdmScalarPropertyAttribute(EntityKeyProperty:=false, IsNullable:=false)>
    <DataMemberAttribute()>
    Public Property Name() As Global.System.String
        Get
            Return _Name
        End Get
        Set
            OnNameChanging(value)
            ReportPropertyChanging("Name")
            _Name = StructuralObject.SetValidValue(value, false)
            ReportPropertyChanged("Name")
            OnNameChanged()
        End Set
    End Property

    Private _Name As Global.System.String
    Private Partial Sub OnNameChanging(value As Global.System.String)
    End Sub

    Private Partial Sub OnNameChanged()
    End Sub

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <EdmScalarPropertyAttribute(EntityKeyProperty:=false, IsNullable:=false)>
    <DataMemberAttribute()>
    Public Property Value() As Global.System.String
        Get
            Return _Value
        End Get
        Set
            OnValueChanging(value)
            ReportPropertyChanging("Value")
            _Value = StructuralObject.SetValidValue(value, false)
            ReportPropertyChanged("Value")
            OnValueChanged()
        End Set
    End Property

    Private _Value As Global.System.String
    Private Partial Sub OnValueChanging(value As Global.System.String)
    End Sub

    Private Partial Sub OnValueChanged()
    End Sub

    #End Region
End Class

''' <summary>
''' No Metadata Documentation available.
''' </summary>
<EdmEntityTypeAttribute(NamespaceName:="TPHModel", Name:="TPHPrimaryData")>
<Serializable()>
<DataContractAttribute(IsReference:=True)>
Public Partial Class TPHPrimaryData
    Inherits TPHDataItem
    #Region "Factory Method"

    ''' <summary>
    ''' Create a new TPHPrimaryData object.
    ''' </summary>
    ''' <param name="index">Initial value of the Index property.</param>
    ''' <param name="name">Initial value of the Name property.</param>
    ''' <param name="value">Initial value of the Value property.</param>
    Public Shared Function CreateTPHPrimaryData(index As Global.System.Int32, name As Global.System.String, value As Global.System.String) As TPHPrimaryData
        Dim tPHPrimaryData as TPHPrimaryData = New TPHPrimaryData
        tPHPrimaryData.Index = index
        tPHPrimaryData.Name = name
        tPHPrimaryData.Value = value
        Return tPHPrimaryData
    End Function

    #End Region
    #Region "Primitive Properties"

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <EdmScalarPropertyAttribute(EntityKeyProperty:=false, IsNullable:=true)>
    <DataMemberAttribute()>
    Public Property PrimaryIdentifier() As Global.System.String
        Get
            Return _PrimaryIdentifier
        End Get
        Set
            OnPrimaryIdentifierChanging(value)
            ReportPropertyChanging("PrimaryIdentifier")
            _PrimaryIdentifier = StructuralObject.SetValidValue(value, true)
            ReportPropertyChanged("PrimaryIdentifier")
            OnPrimaryIdentifierChanged()
        End Set
    End Property

    Private _PrimaryIdentifier As Global.System.String
    Private Partial Sub OnPrimaryIdentifierChanging(value As Global.System.String)
    End Sub

    Private Partial Sub OnPrimaryIdentifierChanged()
    End Sub

    #End Region
    #Region "Navigation Properties"

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <XmlIgnoreAttribute()>
    <SoapIgnoreAttribute()>
    <DataMemberAttribute()>
    <EdmRelationshipNavigationPropertyAttribute("TPHModel", "TPHPrimaryDataTPHSecondaryData", "TPHSecondaryData")>
     Public Property TPHSecondaryDatas() As EntityCollection(Of TPHSecondaryData)
        Get
            Return CType(Me,IEntityWithRelationships).RelationshipManager.GetRelatedCollection(Of TPHSecondaryData)("TPHModel.TPHPrimaryDataTPHSecondaryData", "TPHSecondaryData")
        End Get
        Set
            If (Not value Is Nothing)
                CType(Me, IEntityWithRelationships).RelationshipManager.InitializeRelatedCollection(Of TPHSecondaryData)("TPHModel.TPHPrimaryDataTPHSecondaryData", "TPHSecondaryData", value)
            End If
        End Set
    End Property

    #End Region
End Class

''' <summary>
''' No Metadata Documentation available.
''' </summary>
<EdmEntityTypeAttribute(NamespaceName:="TPHModel", Name:="TPHSecondaryData")>
<Serializable()>
<DataContractAttribute(IsReference:=True)>
Public Partial Class TPHSecondaryData
    Inherits TPHDataItem
    #Region "Factory Method"

    ''' <summary>
    ''' Create a new TPHSecondaryData object.
    ''' </summary>
    ''' <param name="index">Initial value of the Index property.</param>
    ''' <param name="name">Initial value of the Name property.</param>
    ''' <param name="value">Initial value of the Value property.</param>
    Public Shared Function CreateTPHSecondaryData(index As Global.System.Int32, name As Global.System.String, value As Global.System.String) As TPHSecondaryData
        Dim tPHSecondaryData as TPHSecondaryData = New TPHSecondaryData
        tPHSecondaryData.Index = index
        tPHSecondaryData.Name = name
        tPHSecondaryData.Value = value
        Return tPHSecondaryData
    End Function

    #End Region
    #Region "Primitive Properties"

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <EdmScalarPropertyAttribute(EntityKeyProperty:=false, IsNullable:=true)>
    <DataMemberAttribute()>
    Public Property SecondaryIdentifier() As Global.System.String
        Get
            Return _SecondaryIdentifier
        End Get
        Set
            OnSecondaryIdentifierChanging(value)
            ReportPropertyChanging("SecondaryIdentifier")
            _SecondaryIdentifier = StructuralObject.SetValidValue(value, true)
            ReportPropertyChanged("SecondaryIdentifier")
            OnSecondaryIdentifierChanged()
        End Set
    End Property

    Private _SecondaryIdentifier As Global.System.String
    Private Partial Sub OnSecondaryIdentifierChanging(value As Global.System.String)
    End Sub

    Private Partial Sub OnSecondaryIdentifierChanged()
    End Sub

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <EdmScalarPropertyAttribute(EntityKeyProperty:=false, IsNullable:=false)>
    <DataMemberAttribute()>
    Public Property TPHPrimaryDataIndex() As Global.System.Int32
        Get
            Return _TPHPrimaryDataIndex
        End Get
        Set
            OnTPHPrimaryDataIndexChanging(value)
            ReportPropertyChanging("TPHPrimaryDataIndex")
            _TPHPrimaryDataIndex = StructuralObject.SetValidValue(value)
            ReportPropertyChanged("TPHPrimaryDataIndex")
            OnTPHPrimaryDataIndexChanged()
        End Set
    End Property

    Private _TPHPrimaryDataIndex As Global.System.Int32 = -1
    Private Partial Sub OnTPHPrimaryDataIndexChanging(value As Global.System.Int32)
    End Sub

    Private Partial Sub OnTPHPrimaryDataIndexChanged()
    End Sub

    #End Region
    #Region "Navigation Properties"

    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <XmlIgnoreAttribute()>
    <SoapIgnoreAttribute()>
    <DataMemberAttribute()>
    <EdmRelationshipNavigationPropertyAttribute("TPHModel", "TPHPrimaryDataTPHSecondaryData", "TPHPrimaryData")>
    Public Property TPHPrimaryData() As TPHPrimaryData
        Get
            Return CType(Me, IEntityWithRelationships).RelationshipManager.GetRelatedReference(Of TPHPrimaryData)("TPHModel.TPHPrimaryDataTPHSecondaryData", "TPHPrimaryData").Value
        End Get
        Set
            CType(Me, IEntityWithRelationships).RelationshipManager.GetRelatedReference(Of TPHPrimaryData)("TPHModel.TPHPrimaryDataTPHSecondaryData", "TPHPrimaryData").Value = value
        End Set
    End Property
    ''' <summary>
    ''' No Metadata Documentation available.
    ''' </summary>
    <BrowsableAttribute(False)>
    <DataMemberAttribute()>
    Public Property TPHPrimaryDataReference() As EntityReference(Of TPHPrimaryData)
        Get
            Return CType(Me, IEntityWithRelationships).RelationshipManager.GetRelatedReference(Of TPHPrimaryData)("TPHModel.TPHPrimaryDataTPHSecondaryData", "TPHPrimaryData")
        End Get
        Set
            If (Not value Is Nothing)
                CType(Me, IEntityWithRelationships).RelationshipManager.InitializeRelatedReference(Of TPHPrimaryData)("TPHModel.TPHPrimaryDataTPHSecondaryData", "TPHPrimaryData", value)
            End If
        End Set
    End Property

    #End Region
End Class

#End Region

