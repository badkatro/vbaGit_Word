VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVBAGitExtractor 
   Caption         =   "VBAGit Extractor"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   OleObjectBlob   =   "frmVBAGitExtractor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVBAGitExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Const ListboxItemHeight As Integer = 13.8     ' how much do you add to a listbox height to fit a new item?
Const maxLBItems As Integer = 8             ' how many items fit into existing listboxes without expanding it?

Public ctProjectName As String



Public Sub PrintToOutput(StringToPrint As String)


frmVBAGitExtractor.tbOutput.Text = IIf(frmVBAGitExtractor.tbOutput.Text = "", StringToPrint, frmVBAGitExtractor.tbOutput.Text & vbLf & StringToPrint)


End Sub


Private Sub SelectVBAComponents(WhichComponents As String)
' WhichComponents can be: "None", "All", "Modules", "NoModules", "Forms", "NoForms", "Classes" & "NoClasses"
' so as to select or deselect the respective categories respectively

Select Case WhichComponents
    
    Case "All"
        For i = 0 To lbModulesList.ListCount - 1: lbModulesList.Selected(i) = True: Next i
        For i = 0 To lbFormsList.ListCount - 1: lbFormsList.Selected(i) = True: Next i
        For i = 0 To lbClassesList.ListCount - 1: lbClassesList.Selected(i) = True: Next i
        
    Case "None"
        For i = 0 To lbModulesList.ListCount - 1: lbModulesList.Selected(i) = False: Next i
        For i = 0 To lbFormsList.ListCount - 1: lbFormsList.Selected(i) = False: Next i
        For i = 0 To lbClassesList.ListCount - 1: lbClassesList.Selected(i) = False: Next i
        
    Case "Modules"
        For i = 0 To lbModulesList.ListCount - 1: lbModulesList.Selected(i) = True: Next i

    Case "NoModules"
        For i = 0 To lbModulesList.ListCount - 1: lbModulesList.Selected(i) = False: Next i

    Case "Forms"
        For i = 0 To lbFormsList.ListCount - 1: lbFormsList.Selected(i) = True: Next i

    Case "NoForms"
        For i = 0 To lbFormsList.ListCount - 1: lbFormsList.Selected(i) = False: Next i

    Case "Classes"
        For i = 0 To lbClassesList.ListCount - 1: lbClassesList.Selected(i) = True: Next i

    Case "NoClasses"
        For i = 0 To lbClassesList.ListCount - 1: lbClassesList.Selected(i) = False: Next i

End Select


End Sub


Private Sub cmdCancel_Click()

Unload Me

End Sub




Private Sub cmdExtract_Click()


tbOutput.Text = ""
tbOutput.SetFocus


Dim selectedVBComponents

selectedVBComponents = GetSelectedVBAComponents

If selectedVBComponents = "" Then frmVBAGitExtractor.PrintToOutput "No VB component selected, Aborted...": Exit Sub

Call extractVBComponents(ctProjectName, CStr(selectedVBComponents), ctProjectName)

End Sub


Private Function GetSelectedVBAComponents() As String

Dim tmpString

If lbModulesList.ListCount > 0 Then
    
    For i = 0 To lbModulesList.ListCount - 1
        
        If lbModulesList.Selected(i) Then
            
            tmpString = IIf(tmpString = "", lbModulesList.list(i), tmpString & "," & lbModulesList.list(i))
            
        End If
        
    Next i
    
End If

If lbFormsList.ListCount > 0 Then
    
    For i = 0 To lbFormsList.ListCount - 1
        
        If lbFormsList.Selected(i) Then
            
            tmpString = IIf(tmpString = "", lbFormsList.list(i), tmpString & "," & lbFormsList.list(i))
            
        End If
        
    Next i
    
End If

If lbClassesList.ListCount > 0 Then
    
    For i = 0 To lbClassesList.ListCount - 1
        
        If lbClassesList.Selected(i) Then
            
            tmpString = IIf(tmpString = "", lbClassesList.list(i), tmpString & "," & lbClassesList.list(i))
            
        End If
        
    Next i
    
End If

GetSelectedVBAComponents = tmpString


End Function


Private Sub cmdExtractForms_Click()

Call export_AllForms_ofVBProject(ctProjectName)

End Sub


Private Sub cmdExtrGit_Click()


tbOutput.Text = ""
tbOutput.SetFocus   ' will display scroll bar if needed and scroll as needed


Dim selectedVBComponents

selectedVBComponents = GetSelectedVBAComponents


If selectedVBComponents = "" Then frmVBAGitExtractor.PrintToOutput "NO VB component selected, Aborted...": Exit Sub


Call extractVBComponents(ctProjectName, CStr(selectedVBComponents), ctProjectName)

Call Git_Repo(ctProjectName)

End Sub



Private Sub cmdGit_Click()

tbOutput.Text = ""
tbOutput.SetFocus


Call Git_Repo(ctProjectName)

End Sub

Private Sub lbClassesList_Change()

If tbOutput.Text <> "" Then tbOutput.Text = ""

End Sub

Private Sub lbClassesList_Click()

End Sub

Private Sub lblDeselectAll_Click()

Call SelectVBAComponents("None")

End Sub

Private Sub lblDeselectClasses_Click()

Call SelectVBAComponents("NoClasses")

End Sub

Private Sub lblDeselectForms_Click()

Call SelectVBAComponents("NoForms")

End Sub

Private Sub lblDeselectMods_Click()

Call SelectVBAComponents("NoModules")

End Sub

Private Sub lblSelectAll_Click()

Call SelectVBAComponents("All")

End Sub

Private Sub lblSelectClasses_Click()

Call SelectVBAComponents("Classes")

End Sub

Private Sub lblSelectForms_Click()

Call SelectVBAComponents("Forms")

End Sub

Private Sub lblSelectMods_Click()

Call SelectVBAComponents("Modules")

End Sub

Private Sub lbModulesList_Change()

If tbOutput.Text <> "" Then tbOutput.Text = ""

End Sub

Private Sub UserForm_Activate()


If lbFormsList.ListCount = 0 Then
    
    ' We have no forms but also no classes !
    If lbClassesList.ListCount = 0 Then
        
        Call HideFormsGroup
        
        Call HideClassesGroup
        
        Call UnWidenForm(248)

    Else
    
        Call HideFormsGroup
        
        Call NotchClassesGroupLeft
        
        Call UnWidenForm(124)
    
    End If

' second case, forms, we have, but not classes... simpler
ElseIf lbClassesList.ListCount = 0 Then
    
    Call HideClassesGroup
    
    Call UnWidenForm(124)
    
End If

End Sub

Private Sub NotchClassesGroupLeft()

lbClassesList.Left = lbFormsList.Left
        lblClasses.Left = lblForms.Left
        lblSelectClasses.Left = lblSelectForms.Left
        lblDeselectClasses.Left = lblDeselectForms.Left

End Sub

Private Sub HideFormsGroup()

lbFormsList.Visible = False
        lblForms.Visible = False
        lblSelectForms.Visible = False
        lblDeselectForms.Visible = False
        
End Sub

Private Sub HideClassesGroup()

lbClassesList.Visible = False
        lblClasses.Visible = False
        lblSelectClasses.Visible = False
        lblDeselectClasses.Visible = False

End Sub

Private Sub UnWidenForm(Differential As Integer)

frmVBAGitExtractor.Width = frmVBAGitExtractor.Width - Differential
    
cmdExtrGit.Left = cmdExtrGit.Left - Differential
cmdExtract.Left = cmdExtract.Left - Differential
cmdGit.Left = cmdGit.Left - Differential
cmdCancel.Left = cmdCancel.Left - Differential
lblGrayLine.Left = lblGrayLine.Left - Differential
tbOutput.Width = tbOutput.Width - Differential


End Sub

Private Sub UserForm_Initialize()

If Documents.Count > 0 Then
    
    If HaveOpenedTemplate Then
        
        Dim tp As Document
        
        Set tp = GetOpenedTemplate
        
        frmVBAGitExtractor.Caption = frmVBAGitExtractor.Caption & " - " & tp.VBProject.name
        
        ctProjectName = tp.VBProject.name
        
        Call ListVbaComponents(tp)
        
    End If
    
End If


End Sub


Function GetOpenedTemplate() As Document

If Documents.Count > 0 Then

    Dim doc As Document
    
    For Each doc In Documents
        
        If Right(doc.name, 5) = ".dotm" Then
            Set GetOpenedTemplate = doc
            Exit Function
        End If
        
    Next doc

End If

Set GetOpenedTemplate = Nothing


End Function

Function HaveOpenedTemplate() As Boolean


If Documents.Count > 0 Then

    Dim doc As Document
    
    For Each doc In Documents
        
        If Right(doc.name, 5) = ".dotm" Then
            HaveOpenedTemplate = True
            Exit Function
        End If
        
    Next doc

End If

HaveOpenedTemplate = False


End Function


Private Sub ListVbaComponents(Template As Document)


Dim vbaComps As VBIDE.VBComponents

Set vbaComps = Template.VBProject.VBComponents

Dim vbaC As VBComponent

For Each vbaC In vbaComps
    
    If vbaC.Type = vbext_ct_StdModule Then
        
        If lbModulesList.BackColor <> wdColorWhite Then lbModulesList.BackColor = wdColorWhite: lbModulesList.SpecialEffect = fmSpecialEffectSunken
        
        If lbModulesList.ListCount > maxLBItems Then
            
            lbModulesList.Height = lbModulesList.Height + ListboxItemHeight
            
            If lbModulesList.Height > lbFormsList.Height And lbModulesList.Height > lbClassesList.Height Then
                Call EnlargeUserform
            End If
        
        End If
        
        lbModulesList.AddItem (vbaC.name)
    
    ElseIf vbaC.Type = vbext_ct_ClassModule Then
        
        If lbClassesList.BackColor <> wdColorWhite Then lbClassesList.BackColor = wdColorWhite: lbClassesList.SpecialEffect = fmSpecialEffectSunken
        
        If lbClassesList.ListCount > maxLBItems Then
            
            lbClassesList.Height = lbClassesList.Height + ListboxItemHeight
            
            If lbClassesList.Height > lbModulesList.Height And lbClassesList.Height > lbFormsList.Height Then
        
                Call EnlargeUserform
            
            End If
        
        End If
        
        lbClassesList.AddItem (vbaC.name)
    
    ElseIf vbaC.Type = vbext_ct_MSForm Then
        
        If lbFormsList.BackColor <> wdColorWhite Then lbFormsList.BackColor = wdColorWhite: lbFormsList.SpecialEffect = fmSpecialEffectSunken
        
        If lbFormsList.ListCount > maxLBItems Then
            
            lbFormsList.Height = lbFormsList.Height + ListboxItemHeight
            
            If lbFormsList.Height > lbModulesList.Height And lbFormsList.Height > lbClassesList.Height Then
                
                Call EnlargeUserform
                
            End If
            
        End If
        
        lbFormsList.AddItem (vbaC.name)
    
    End If
    
Next vbaC

End Sub


Private Sub EnlargeUserform()

frmVBAGitExtractor.Height = frmVBAGitExtractor.Height + ListboxItemHeight
'cmdExtrGit.Top = cmdExtrGit.Top + ListboxItemHeight: cmdExtract.Top = cmdExtract.Top + ListboxItemHeight
'cmdGit.Top = cmdGit.Top + ListboxItemHeight
cmdCancel.Top = cmdCancel.Top + ListboxItemHeight
tbOutput.Top = tbOutput.Top + ListboxItemHeight
lblOutput.Top = lblOutput.Top + ListboxItemHeight
lblGrayBack.Height = lblGrayBack.Height + ListboxItemHeight
lblGrayLine.Height = lblGrayLine.Height + ListboxItemHeight

frmVBAGitExtractor.Repaint

End Sub



