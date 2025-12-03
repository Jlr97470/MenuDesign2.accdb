Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Function ShowSubMenuElement()
'On Error Resume Next

'Reset all button and text box before do anything
        For j = 1 To 5
            Me("B" & j).Visible = False
            Me("T" & j).Value = Null
        Next

        Me.Fram2.SourceObject = ""
 '------------------------------------------------

        Lb0.Caption = Me.ActiveControl.Caption
        Fram1.Visible = True
        Fram1.SourceObject = "DefaultPageF"
        Fram2.Visible = False

       Dim MenuSelectedID As Double
        MenuSelectedID = CDbl(Nz(Mid(Me.ActiveControl.Name, 2), 0)) ' Main menu button position number detection

       Dim MenuSelectedElementCountInTable As Integer 'numer of sub menu element detected in the table
        MenuSelectedElementCountInTable = Nz(DCount("id", "MenuDataTbl", "MenuId=" & MenuSelectedID), 0)

       Dim FirstIDofTheMenuElement As Double 'Detect the first ID
        FirstIDofTheMenuElement = CDbl(Nz(DLookup("ID", "MenuDataTbl", "menuid =" & MenuSelectedID), 0))

            'MsgBox FirstIDofTheMenuElement
            If FirstIDofTheMenuElement = 0 Then Exit Function



    For i = 1 To MenuSelectedElementCountInTable
            'set the sub element caption
            Me("B" & i).Caption = DLookup("ButtonCaptionStr", "MenuDataTbl", "menuid =" & MenuSelectedID & " and ID =" & FirstIDofTheMenuElement)

            'set the form name in to the textbox below the sub element button
            Me("T" & i).Value = DLookup("FormNameStr", "MenuDataTbl", "menuid =" & MenuSelectedID & " and ID =" & FirstIDofTheMenuElement)
              'Me("T" & i).Value = DLookup("FormNameStr", "MenuDataTbl", "menuid =" & MenuSelectedID & " and ID =" & FirstIDofTheMenuElement)


            'set next row id information
            FirstIDofTheMenuElement = CDbl(Nz(DLookup("ID", "MenuDataTbl", "menuid =" & MenuSelectedID & " and ID >" & FirstIDofTheMenuElement), 0))

            'visibility
            Me("B" & i).Visible = True

    Next

End Function
Private Function TriggerSubelementCmd()
On Error Resume Next

    Me.Fram2.SourceObject = Me("T" & Mid(Me.ActiveControl.Name, 2)).Value
    Fram2.Visible = True


End Function


Private Sub M1_Click()
 'On Error Resume Next
    Call ShowSubMenuElement
        Fram1.SourceObject = "itemf"

End Sub

Private Sub M2_Click()
On Error Resume Next
   Call ShowSubMenuElement
    Fram1.SourceObject = "salef"

End Sub

Private Sub M3_Click()
On Error Resume Next
   Call ShowSubMenuElement
    Fram1.SourceObject = "supplierf"
End Sub
