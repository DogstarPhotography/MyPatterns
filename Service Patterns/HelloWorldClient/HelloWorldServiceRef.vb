﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.225
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On



<System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0"),  _
 System.ServiceModel.ServiceContractAttribute(ConfigurationName:="IHelloWorldService")>  _
Public Interface IHelloWorldService
    
    <System.ServiceModel.OperationContractAttribute(Action:="http://tempuri.org/IHelloWorldService/GetMessage", ReplyAction:="http://tempuri.org/IHelloWorldService/GetMessageResponse")>  _
    Function GetMessage(ByVal name As String) As String
End Interface

<System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")>  _
Public Interface IHelloWorldServiceChannel
    Inherits IHelloWorldService, System.ServiceModel.IClientChannel
End Interface

<System.Diagnostics.DebuggerStepThroughAttribute(),  _
 System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")>  _
Partial Public Class HelloWorldServiceClient
    Inherits System.ServiceModel.ClientBase(Of IHelloWorldService)
    Implements IHelloWorldService
    
    Public Sub New()
        MyBase.New
    End Sub
    
    Public Sub New(ByVal endpointConfigurationName As String)
        MyBase.New(endpointConfigurationName)
    End Sub
    
    Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As String)
        MyBase.New(endpointConfigurationName, remoteAddress)
    End Sub
    
    Public Sub New(ByVal endpointConfigurationName As String, ByVal remoteAddress As System.ServiceModel.EndpointAddress)
        MyBase.New(endpointConfigurationName, remoteAddress)
    End Sub
    
    Public Sub New(ByVal binding As System.ServiceModel.Channels.Binding, ByVal remoteAddress As System.ServiceModel.EndpointAddress)
        MyBase.New(binding, remoteAddress)
    End Sub
    
    Public Function GetMessage(ByVal name As String) As String Implements IHelloWorldService.GetMessage
        Return MyBase.Channel.GetMessage(name)
    End Function
End Class
