Attribute VB_Name = "MaintainDremioToken"
Option Explicit

Public prop As DocumentProperty
Public propertyExists As Boolean

Sub AddDocumentVariable(dt As String)
    Dim docType As Office.MsoDocProperties
    propertyExists = False
    docType = msoPropertyTypeString
    
    For Each prop In ActiveWorkbook.CustomDocumentProperties
        If prop.Name = "dremioToken" Then
            propertyExists = True
            ActiveWorkbook.CustomDocumentProperties("dremioToken").Delete
            Exit For
        End If
    Next prop
        
    ActiveWorkbook.CustomDocumentProperties.Add Name:="dremioToken", LinkToContent:=False, Type:=docType, Value:=dt
End Sub

Sub UseDocumentVariable()
    'Dim dremioToken As String
    propertyExists = False
    
    For Each prop In ActiveWorkbook.CustomDocumentProperties
        If prop.Name = "dremioToken" Then
            propertyExists = True
            dremioToken = prop.Value
            Exit For
        End If
    Next prop
    
    If propertyExists <> True Then
        Connect
    End If
        
End Sub
