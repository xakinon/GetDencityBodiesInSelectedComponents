Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
'Imports NXOpenUI

Module NXJournal

Sub Main()

    Dim theSession As Session = Session.GetSession()
    Dim workPart As Part = theSession.Parts.Work
    Dim displayPart As Part = theSession.Parts.Display
    Dim lw As ListingWindow = theSession.ListingWindow
    Dim mySelectedComponents As New List(Of Assemblies.Component)
    Dim theUI As UI = UI.GetUI()

    ' コンポーネントの選択
    If SelectObjects("コンポーネントを選択してください", mySelectedComponents) <> Selection.Response.Ok Then
        return
    End If

    lw.open()

    For Each theComponent As Assemblies.Component in mySelectedComponents
        Dim myPart As Part = CType(theComponent.Prototype.OwningPart, Part)
        For Each myBody As Body In myPart.Bodies
            If myBody.IsSolidBody() Then
                lw.writeline( myBody.Density )
            End If
        Next
    Next

End Sub

Function SelectObjects(prompt As String, ByRef dispObj As List(Of Assemblies.Component)) As Selection.Response

    Dim selObj As NXObject()
    Dim theUI As UI = UI.GetUI
    Dim title As String = "Select objects"
    Dim includeFeatures As Boolean = False
    Dim keepHighlighted As Boolean = False
    Dim selAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific
    Dim scope As Selection.SelectionScope = Selection.SelectionScope.AnyInAssembly
    Dim selectionMask_array(0) As Selection.MaskTriple

    With selectionMask_array(0)
        .Type = UF.UFConstants.UF_component_type
        .Subtype = UF.UFConstants.UF_component_subtype
    End With

    Dim resp As Selection.Response = theUI.SelectionManager.SelectObjects(prompt, title, scope, selAction, includeFeatures, keepHighlighted, selectionMask_array, selObj)

    'If resp = Selection.Response.ObjectSelected Or
    '        resp = Selection.Response.ObjectSelectedByName Or
    '        resp = Selection.Response.Ok Then
    If resp = Selection.Response.Ok Then
        For Each item As NXObject In selObj
            dispObj.Add(CType(item, DisplayableObject))
        Next
        Return Selection.Response.Ok
    Else
        Return Selection.Response.Cancel
    End If

End Function

End Module


