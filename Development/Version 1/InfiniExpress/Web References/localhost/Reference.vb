﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.0.3705.288
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'This source code was auto-generated by Microsoft.VSDesigner, Version 1.0.3705.288.
'
Namespace localhost
    
    '<remarks/>
    <System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="InfiniServiceSoap", [Namespace]:="http://www.accountscentre.com/")>  _
    Public Class InfiniService
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        '<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = "http://localhost/InfiniWebService/Default.asmx"
        End Sub
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.accountscentre.com/CustomerIdValidation", RequestNamespace:="http://www.accountscentre.com/", ResponseNamespace:="http://www.accountscentre.com/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function CustomerIdValidation(ByVal CustomerId As Integer, ByVal Password As String) As String
            Dim results() As Object = Me.Invoke("CustomerIdValidation", New Object() {CustomerId, Password})
            Return CType(results(0),String)
        End Function
        
        '<remarks/>
        Public Function BeginCustomerIdValidation(ByVal CustomerId As Integer, ByVal Password As String, ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("CustomerIdValidation", New Object() {CustomerId, Password}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndCustomerIdValidation(ByVal asyncResult As System.IAsyncResult) As String
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),String)
        End Function
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.accountscentre.com/SynchronizeFromWeb", RequestNamespace:="http://www.accountscentre.com/", ResponseNamespace:="http://www.accountscentre.com/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function SynchronizeFromWeb(ByVal DTds As System.Data.DataSet) As Boolean
            Dim results() As Object = Me.Invoke("SynchronizeFromWeb", New Object() {DTds})
            Return CType(results(0),Boolean)
        End Function
        
        '<remarks/>
        Public Function BeginSynchronizeFromWeb(ByVal DTds As System.Data.DataSet, ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("SynchronizeFromWeb", New Object() {DTds}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndSynchronizeFromWeb(ByVal asyncResult As System.IAsyncResult) As Boolean
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),Boolean)
        End Function
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.accountscentre.com/SynchronizeFromDeskTop", RequestNamespace:="http://www.accountscentre.com/", ResponseNamespace:="http://www.accountscentre.com/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function SynchronizeFromDeskTop(ByVal WSds As System.Data.DataSet) As System.Data.DataSet
            Dim results() As Object = Me.Invoke("SynchronizeFromDeskTop", New Object() {WSds})
            Return CType(results(0),System.Data.DataSet)
        End Function
        
        '<remarks/>
        Public Function BeginSynchronizeFromDeskTop(ByVal WSds As System.Data.DataSet, ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("SynchronizeFromDeskTop", New Object() {WSds}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndSynchronizeFromDeskTop(ByVal asyncResult As System.IAsyncResult) As System.Data.DataSet
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),System.Data.DataSet)
        End Function
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.accountscentre.com/UpdateSynchronizeFlag", RequestNamespace:="http://www.accountscentre.com/", ResponseNamespace:="http://www.accountscentre.com/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function UpdateSynchronizeFlag(ByVal WSds As System.Data.DataSet) As Boolean
            Dim results() As Object = Me.Invoke("UpdateSynchronizeFlag", New Object() {WSds})
            Return CType(results(0),Boolean)
        End Function
        
        '<remarks/>
        Public Function BeginUpdateSynchronizeFlag(ByVal WSds As System.Data.DataSet, ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("UpdateSynchronizeFlag", New Object() {WSds}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndUpdateSynchronizeFlag(ByVal asyncResult As System.IAsyncResult) As Boolean
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),Boolean)
        End Function
    End Class
End Namespace
