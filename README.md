# Transfer-Script
A macro enabled template that will pull a table of steps from a reference document and then add the necessary information based on selections from the userform.

Private Sub LAS7transfer()

    Dim Criticalitylevel_7 As Range
    Dim projectname_7 As Range
    Dim projectname_1_7 As Range
    Dim address_7 As Range
    Dim Title_7 As Range
    Dim PG_7 As Range
    Dim pg_1_7 As Range
    Dim PG_2_7 As Range
    Dim PG_3_7 As Range
    Dim PG_4_7 As Range
    Dim PG_5_7 As Range
    Dim PG_6_7 As Range
    Dim PG_7_7 As Range
    Dim USB_2_7 As Range
    Dim USB_3_7 As Range
    Dim USB_4_7 As Range
    Dim USB_5_7 As Range
    Dim EquipmentID_7 As Range
    Dim PGUSB_7 As Range
    Dim PGUSB_3_7 As Range
    Dim S_7 As Range
    Dim SPA_7 As Range
    Dim spa_3_7 As Range
    Dim SPA_2_7 As Range
    Dim SPB_7 As Range
    Dim SPB_2_7 As Range
    Dim SPB_3_7 As Range
    Dim SPB_4_7 As Range
    Dim spc_1_7 As Range
    Dim SPC_3_7 As Range
    Dim SPC_4_7 As Range
    Dim ISX_7 As Range
    Dim ISX_2_7 As Range
    Dim UPS_7 As Range
    Dim UPS_1_7 As Range
    Dim UPS_2_7 As Range
    Dim UPS_3_7 As Range
    Dim UPS_4_7 As Range
    Dim UPS_5_7 As Range
    Dim UPS_6_7 As Range
    Dim UPS_7_7 As Range
    Dim UPS_8_7 As Range
    Dim UPS_9_7 As Range
    Dim UPS_10_7 As Range
    Dim UPS_11_7 As Range
    Dim UPS_12_7 As Range
    Dim BuildingName_7 As Range
    Dim footerSite_7 As Range
    Dim footerSite_7_1 As Range
    Dim targettable As table
    
    
    
Set targetdoc = Documents.Open(FileName:=Environ("USERPROFILE") & "\Desktop\" & ProjectName.Text & " " & "transfer script" & " " & Month(Now) _
    & "." & Day(Now) & "." & Year(Now) & ".docx")
    
setpduLAS7
If cboSite.value = "LAS 7" And cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Then
On Error GoTo Errorhandler_8
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Single UPS Annual or Corrective.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_1_7 = Nothing
    Set UPS_2_7 = Nothing
    Set UPS_3_7 = Nothing
    Set UPS_4_7 = Nothing
    Set UPS_5_7 = Nothing
    Set UPS_6_7 = Nothing
    Set UPS_7_7 = Nothing
    Set UPS_8_7 = Nothing
    Set UPS_9_7 = Nothing
    Set UPS_10_7 = Nothing
    Set UPS_11_7 = Nothing
    Set UPS_12_7 = Nothing
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = Nothing
    EquipmentID_7.Text = cboEquipmentID.value
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    generatorselectionLAS7
    
Errorhandler_8:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Single UPS Annual or Corrective.docx")
            End Select
Resume Next

ElseIf cboSite.value = "LAS 7" And cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Corrective Maintenance" Then
On Error GoTo Errorhandler_9
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Single UPS Annual or Corrective.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_1_7 = Nothing
    Set UPS_2_7 = Nothing
    Set UPS_3_7 = Nothing
    Set UPS_4_7 = Nothing
    Set UPS_5_7 = Nothing
    Set UPS_6_7 = Nothing
    Set UPS_7_7 = Nothing
    Set UPS_8_7 = Nothing
    Set UPS_9_7 = Nothing
    Set UPS_10_7 = Nothing
    Set UPS_11_7 = Nothing
    Set UPS_12_7 = Nothing
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = Nothing
    EquipmentID_7.Text = cboEquipmentID.value
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    generatorselectionLAS7
    
Errorhandler_9:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Single UPS Annual or Corrective.docx")
            End Select
Resume Next

ElseIf cboSite.value = "LAS 7" And cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal" Then
On Error GoTo Errorhandler_10
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Single UPS Annual w Cal.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_1_7 = Nothing
    Set UPS_2_7 = Nothing
    Set UPS_3_7 = Nothing
    Set UPS_4_7 = Nothing
    Set UPS_5_7 = Nothing
    Set UPS_6_7 = Nothing
    Set UPS_7_7 = Nothing
    Set UPS_8_7 = Nothing
    Set UPS_9_7 = Nothing
    Set UPS_10_7 = Nothing
    Set UPS_11_7 = Nothing
    Set UPS_12_7 = Nothing
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = Nothing
    EquipmentID_7.Text = cboEquipmentID.value
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    generatorselectionLAS7
    
Errorhandler_10:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Single UPS Annual w Cal.docx")
            End Select
Resume Next
    
ElseIf cboSite.value = "LAS 7" And cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Then
On Error GoTo Errorhandler_11
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Single UPS Annual w Cal and Depl.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc.Tables
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_1_7 = Nothing
    Set UPS_2_7 = Nothing
    Set UPS_3_7 = Nothing
    Set UPS_4_7 = Nothing
    Set UPS_5_7 = Nothing
    Set UPS_6_7 = Nothing
    Set UPS_7_7 = Nothing
    Set UPS_8_7 = Nothing
    Set UPS_9_7 = Nothing
    Set UPS_10_7 = Nothing
    Set UPS_11_7 = Nothing
    Set UPS_12_7 = Nothing
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = Nothing
    EquipmentID_7.Text = cboEquipmentID.value
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    generatorselectionLAS7
    
Errorhandler_11:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Single UPS Annual w Cal and Depl.docx")
            End Select
Resume Next
    
    
ElseIf cboSite.value = "LAS 7" And cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Then
On Error GoTo Errorhandler_12
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Multiple UPS Annual or Corrective.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc.Tables
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_1_7 = ActiveDocument.Bookmarks("tUPS_1").Range
    Set UPS_2_7 = ActiveDocument.Bookmarks("tUPS_2").Range
    Set UPS_3_7 = ActiveDocument.Bookmarks("tUPS_3").Range
    Set UPS_4_7 = ActiveDocument.Bookmarks("tUPS_4").Range
    Set UPS_5_7 = ActiveDocument.Bookmarks("tUPS_5").Range
    Set UPS_6_7 = ActiveDocument.Bookmarks("tUPS_6").Range
    Set UPS_7_7 = ActiveDocument.Bookmarks("tUPS_7").Range
    Set UPS_8_7 = ActiveDocument.Bookmarks("tUPS_8").Range
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
    EquipmentID_7.Text = cboEquipmentID.value
    equipmentID_2_7.Text = "and " & cboEquipmentID_2.value
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    UPS_1_7.Text = cboEquipmentID.value
    UPS_2_7.Text = cboEquipmentID_2.value
    UPS_3_7.Text = cboEquipmentID.value
    UPS_4_7.Text = cboEquipmentID.value
    UPS_5_7.Text = cboEquipmentID_2.value
    UPS_6_7.Text = cboEquipmentID.value
    UPS_7_7.Text = cboEquipmentID_2.value
    UPS_8_7.Text = cboEquipmentID_2.value
    generatorselectionLAS7
    
Errorhandler_12:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Multiple UPS Annual or Corrective.docx")
            End Select
Resume Next
    
ElseIf cboSite.value = "LAS 7" And cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance" Then
On Error GoTo Errorhandler_13
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Multiple UPS Annual or Corrective.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc.Tables
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_1_7 = ActiveDocument.Bookmarks("tUPS_1").Range
    Set UPS_2_7 = ActiveDocument.Bookmarks("tUPS_2").Range
    Set UPS_3_7 = ActiveDocument.Bookmarks("tUPS_3").Range
    Set UPS_4_7 = ActiveDocument.Bookmarks("tUPS_4").Range
    Set UPS_5_7 = ActiveDocument.Bookmarks("tUPS_5").Range
    Set UPS_6_7 = ActiveDocument.Bookmarks("tUPS_6").Range
    Set UPS_7_7 = ActiveDocument.Bookmarks("tUPS_7").Range
    Set UPS_8_7 = ActiveDocument.Bookmarks("tUPS_8").Range
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
    EquipmentID_7.Text = cboEquipmentID.value
    equipmentID_2_7.Text = "and " & cboEquipmentID_2.value
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    UPS_1_7.Text = cboEquipmentID.value
    UPS_2_7.Text = cboEquipmentID_2.value
    UPS_3_7.Text = cboEquipmentID.value
    UPS_4_7.Text = cboEquipmentID.value
    UPS_5_7.Text = cboEquipmentID_2.value
    UPS_6_7.Text = cboEquipmentID.value
    UPS_7_7.Text = cboEquipmentID_2.value
    UPS_8_7.Text = cboEquipmentID_2.value
    generatorselectionLAS7

Errorhandler_13:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Multiple UPS Annual or Corrective.docx")
            End Select
Resume Next


ElseIf cboSite.value = "LAS 7" And cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal" Then
On Error GoTo Errorhandler_14
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Multiple UPS Annual w Cal.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc.Tables
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_7 = ActiveDocument.Bookmarks("tUPS").Range
    Set UPS_1_7 = ActiveDocument.Bookmarks("tUPS_1").Range
    Set UPS_2_7 = ActiveDocument.Bookmarks("tUPS_2").Range
    Set UPS_3_7 = ActiveDocument.Bookmarks("tUPS_3").Range
    Set UPS_4_7 = ActiveDocument.Bookmarks("tUPS_4").Range
    Set UPS_5_7 = ActiveDocument.Bookmarks("tUPS_5").Range
    Set UPS_6_7 = ActiveDocument.Bookmarks("tUPS_6").Range
    Set UPS_7_7 = ActiveDocument.Bookmarks("tUPS_7").Range
    Set UPS_8_7 = ActiveDocument.Bookmarks("tUPS_8").Range
    Set UPS_9_7 = ActiveDocument.Bookmarks("tUPS_9").Range
    Set UPS_10_7 = ActiveDocument.Bookmarks("tUPS_10").Range
    Set UPS_11_7 = ActiveDocument.Bookmarks("tUPS_11").Range
    Set UPS_12_7 = ActiveDocument.Bookmarks("tUPS_12").Range
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    EquipmentID_7.Text = cboEquipmentID.value
    equipmentID_2_7.Text = "and " & cboEquipmentID_2.value
    UPS_7.Text = cboEquipmentID.value
    UPS_1_7.Text = cboEquipmentID.value
    UPS_2_7.Text = cboEquipmentID_2.value
    UPS_3_7.Text = cboEquipmentID.value
    UPS_4_7.Text = cboEquipmentID.value
    UPS_5_7.Text = cboEquipmentID_2.value
    UPS_6_7.Text = cboEquipmentID.value
    UPS_7_7.Text = cboEquipmentID_2.value
    UPS_8_7.Text = cboEquipmentID_2.value
    UPS_9_7.Text = cboEquipmentID.value
    UPS_10_7.Text = cboEquipmentID_2.value
    UPS_11_7.Text = cboEquipmentID.value
    UPS_12_7.Text = cboEquipmentID_2.value
    generatorselectionLAS7
    
Errorhandler_14:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Multiple UPS Annual w Cal.docx")
            End Select
Resume Next

ElseIf cboSite.value = "LAS 7" And cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Then
On Error GoTo Errorhandler_15
Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\LAS7 Multiple UPS Annual w Cal and Depl.docx")
Set targettable = referencedoc.Tables(1)
 For Each targettable In referencedoc.Tables
    targettable.Range.Select
    Debug.Print targettable.Title
    Selection.Copy
    referencedoc.Close
    targetdoc.Activate
    Set startpoint = targetdoc.Paragraphs(146).Range
    startpoint.Paste
  Next targettable
    Set ISX_7 = ActiveDocument.Bookmarks("tISX").Range
    Set ISX_2_7 = ActiveDocument.Bookmarks("tISX_2").Range
    Set projectname_7 = ActiveDocument.Bookmarks("tProjectName").Range
    Set projectname_1_7 = ActiveDocument.Bookmarks("tProjectName_1").Range
    projectname_7.Text = Me.ProjectName.value
    projectname_1_7.Text = Me.ProjectName.value
    Set BuildingName_7 = ActiveDocument.Bookmarks("tbuildingName").Range
    BuildingName_7.Text = "LAS 7"
    Set SPB_4_7 = ActiveDocument.Bookmarks("tSPB_4").Range
    Set UPS_7 = ActiveDocument.Bookmarks("tUPS").Range
    Set UPS_1_7 = ActiveDocument.Bookmarks("tUPS_1").Range
    Set UPS_2_7 = ActiveDocument.Bookmarks("tUPS_2").Range
    Set UPS_3_7 = ActiveDocument.Bookmarks("tUPS_3").Range
    Set UPS_4_7 = ActiveDocument.Bookmarks("tUPS_4").Range
    Set UPS_5_7 = ActiveDocument.Bookmarks("tUPS_5").Range
    Set UPS_6_7 = ActiveDocument.Bookmarks("tUPS_6").Range
    Set UPS_7_7 = ActiveDocument.Bookmarks("tUPS_7").Range
    Set UPS_8_7 = ActiveDocument.Bookmarks("tUPS_8").Range
    Set UPS_9_7 = ActiveDocument.Bookmarks("tUPS_9").Range
    Set UPS_10_7 = ActiveDocument.Bookmarks("tUPS_10").Range
    Set UPS_11_7 = ActiveDocument.Bookmarks("tUPS_11").Range
    Set UPS_12_7 = ActiveDocument.Bookmarks("tUPS_12").Range
    Set EquipmentID_7 = ActiveDocument.Bookmarks("tEquipmentID").Range
    Set equipmentID_2_7 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
    EquipmentID_7.Text = cboEquipmentID.value
    equipmentID_2_7.Text = "and " & cboEquipmentID_2.value
    Set Title_7 = ActiveDocument.Bookmarks("ttitle").Range
    Set address_7 = ActiveDocument.Bookmarks("taddress").Range
    address_7.Text = "7135 S. Decatur Blvd.," & vbCr & "Las Vegas, NV 89118"
    UPS_7.Text = cboEquipmentID.value
    UPS_1_7.Text = cboEquipmentID.value
    UPS_2_7.Text = cboEquipmentID_2.value
    UPS_3_7.Text = cboEquipmentID.value
    UPS_4_7.Text = cboEquipmentID.value
    UPS_5_7.Text = cboEquipmentID_2.value
    UPS_6_7.Text = cboEquipmentID.value
    UPS_7_7.Text = cboEquipmentID_2.value
    UPS_8_7.Text = cboEquipmentID_2.value
    UPS_9_7.Text = cboEquipmentID.value
    UPS_10_7.Text = cboEquipmentID_2.value
    UPS_11_7.Text = cboEquipmentID.value
    UPS_12_7.Text = cboEquipmentID_2.value
    generatorselectionLAS7
    
Errorhandler_15:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/LAS7 Multiple UPS Annual w Cal and Depl.docx")
            End Select
Resume Next

Else

MsgBox "Your selection does not make sense. Please correct your selections."

End If



    
    'Disabling and Enabling ISX Statements if Full Transfer
    

    
    If cboEquipmentID.value = "UPS 1A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 2A" Then
           
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 7A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 8A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 14A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 25A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 26A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 31A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 32A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 37A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 43A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    ElseIf cboEquipmentID.value = "UPS 44A" Then
        
        Title_7.Text = cboSite.value & " Full (A) Building Transfer Script"
        ISX_7.Text = "All UPS ISX"
        ISX_2_7.Text = "All UPS ISX"
        SPB_4_7.Text = "Full"
        
    End If
    
    'Disabling and Enabling ISX Statements if Blue Transfer
    
    If cboEquipmentID.value = "UPS 3B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S2PB"
        
    ElseIf cboEquipmentID.value = "UPS 4B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S2PB"
        
    ElseIf cboEquipmentID.value = "UPS 9B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S1PB"
        
    ElseIf cboEquipmentID.value = "UPS 10B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S1PB"
        
    ElseIf cboEquipmentID.value = "UPS 16B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S3PB"
        
    ElseIf cboEquipmentID.value = "UPS 27B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S5PB"
        
    ElseIf cboEquipmentID.value = "UPS 28B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S5PB"
        
    ElseIf cboEquipmentID.value = "UPS 33B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S6PB"

    ElseIf cboEquipmentID.value = "UPS 34B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S6PB"
        
    ElseIf cboEquipmentID.value = "UPS 39B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S7PB"
        
    ElseIf cboEquipmentID.value = "UPS 45B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S8PB"
        
    ElseIf cboEquipmentID.value = "UPS 46B" Then
        
        Title_7.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX_7.Text = "Both Blue UPS ISX"
        ISX_2_7.Text = "Both Blue UPS ISX"
        SPB_4_7.Text = "S8PB"
        
    End If
       
    'Disabling and Enabling ISX Statements if Grey Transfer
    
     If cboEquipmentID.value = "UPS 5C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S1PC"
        
    ElseIf cboEquipmentID.value = "UPS 6C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S1PC"

    ElseIf cboEquipmentID.value = "UPS 11C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S2PC"
        
    ElseIf cboEquipmentID.value = "UPS 12C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S2PC"
        
    ElseIf cboEquipmentID.value = "UPS 18C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S3PC"
        
    ElseIf cboEquipmentID.value = "UPS 29C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S5PC"
        
    ElseIf cboEquipmentID.value = "UPS 30C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S5PC"
        
    ElseIf cboEquipmentID.value = "UPS 35C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S6PC"
        
    ElseIf cboEquipmentID.value = "UPS 36C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S6PC"
        
    ElseIf cboEquipmentID.value = "UPS 41C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S7PC"
        
    ElseIf cboEquipmentID.value = "UPS 47C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S8PC"
         
    ElseIf cboEquipmentID.value = "UPS 48C" Then
        
        Title_7.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX_7.Text = "Both Grey UPS ISX"
        ISX_2_7.Text = "Both Grey UPS ISX"
        SPB_4_7.Text = "S8PC"
        
    End If
    

'If Power System 1 is selected

    If cboPS.value = "Power System 1 " Then
        Set PG_7 = ActiveDocument.Bookmarks("tPG").Range
        PG_7.Text = "PG1"
        Set pg_1_7 = ActiveDocument.Bookmarks("tPG_1").Range
        pg_1_7.Text = "PG1"
        Set PG_2_7 = ActiveDocument.Bookmarks("tPG_2").Range
        PG_2_7.Text = "PG1"
        Set PG_3_7 = ActiveDocument.Bookmarks("tPG_3").Range
        PG_3_7.Text = "PG1"
        Set PG_4_7 = ActiveDocument.Bookmarks("tPG_4").Range
        PG_4_7.Text = "PG1"
        Set PG_5_7 = ActiveDocument.Bookmarks("tPG_5").Range
        PG_5_7.Text = "PG1"
        Set PG_6_7 = ActiveDocument.Bookmarks("tPG_6").Range
        PG_6_7.Text = "PG1"
        Set PG_7_7 = ActiveDocument.Bookmarks("tPG_7").Range
        PG_7_7.Text = "PG1"
        Set USB_2_7 = ActiveDocument.Bookmarks("tUSB").Range
        USB_2_7.Text = "USB1"
        Set USB_3_7 = ActiveDocument.Bookmarks("tUSB_4").Range
        USB_3_7.Text = "USB1"
        Set USB_4_7 = ActiveDocument.Bookmarks("tUSB_1").Range
        USB_4_7.Text = "USB1"
        Set USB_5_7 = ActiveDocument.Bookmarks("tUSB_2").Range
        USB_5_7.Text = "USB1"
        Set PGUSB_7 = ActiveDocument.Bookmarks("tPGUSB").Range
        PGUSB_7.Text = "PG1-USB1"
        Set PGUSB_3_7 = ActiveDocument.Bookmarks("tPGUSB_2").Range
        PGUSB_3_7.Text = "PG1-USB1"
        Set S_7 = ActiveDocument.Bookmarks("tS").Range
        S_7.Text = "S1"
        Set SPA_7 = ActiveDocument.Bookmarks("tSPA").Range
        SPA_7.Text = "S1PA"
        Set spa_3_7 = ActiveDocument.Bookmarks("tspa_1").Range
        spa_3_7.Text = "S1PA"
        Set SPA_2_7 = ActiveDocument.Bookmarks("tSPA_2").Range
        SPA_2_7.Text = "S1PA"
        Set SPB_7 = ActiveDocument.Bookmarks("tSPB").Range
        SPB_7.Text = "S2PB"
        Set SPB_2_7 = ActiveDocument.Bookmarks("tSPB_2").Range
        SPB_2_7.Text = "S2PB"
        Set SPB_3_7 = ActiveDocument.Bookmarks("tSPB_3").Range
        SPB_3_7.Text = "S2PB"
        Set spc_1_7 = ActiveDocument.Bookmarks("tspc_1").Range
        spc_1_7.Text = "S1PC"
        Set SPC_3_7 = ActiveDocument.Bookmarks("tSPC_3").Range
        SPC_3_7.Text = "S1PC"
        Set SPC_4_7 = ActiveDocument.Bookmarks("tSPC_4").Range
        SPC_4_7.Text = "S1PC"
        
    End If

'If Power System 2 is selected

    If cboPS.value = "Power System 2 " Then
        Set PG_7 = ActiveDocument.Bookmarks("tPG").Range
        PG_7.Text = "PG2"
        Set pg_1_7 = ActiveDocument.Bookmarks("tPG_1").Range
        pg_1_7.Text = "PG2"
        Set PG_2_7 = ActiveDocument.Bookmarks("tPG_2").Range
        PG_2_7.Text = "PG2"
        Set PG_3_7 = ActiveDocument.Bookmarks("tPG_3").Range
        PG_3_7.Text = "PG2"
        Set PG_4_7 = ActiveDocument.Bookmarks("tPG_4").Range
        PG_4_7.Text = "PG2"
        Set PG_5_7 = ActiveDocument.Bookmarks("tPG_5").Range
        PG_5_7.Text = "PG2"
        Set PG_6_7 = ActiveDocument.Bookmarks("tPG_6").Range
        PG_6_7.Text = "PG2"
        Set PG_7_7 = ActiveDocument.Bookmarks("tPG_7").Range
        PG_7_7.Text = "PG2"
        Set USB_2_7 = ActiveDocument.Bookmarks("tUSB").Range
        USB_2_7.Text = "USB2"
        Set USB_3_7 = ActiveDocument.Bookmarks("tUSB_4").Range
        USB_3_7.Text = "USB2"
        Set USB_4_7 = ActiveDocument.Bookmarks("tUSB_1").Range
        USB_4_7.Text = "USB2"
        Set USB_5_7 = ActiveDocument.Bookmarks("tUSB_2").Range
        USB_5_7.Text = "USB2"
        Set PGUSB_7 = ActiveDocument.Bookmarks("tPGUSB").Range
        PGUSB_7.Text = "PG2-USB2"
        Set PGUSB_3_7 = ActiveDocument.Bookmarks("tPGUSB_2").Range
        PGUSB_3_7.Text = "PG2-USB2"
        Set S_7 = ActiveDocument.Bookmarks("tS").Range
        S_7.Text = "S2"
        Set SPA_7 = ActiveDocument.Bookmarks("tSPA").Range
        SPA_7.Text = "S2PA"
        Set spa_3_7 = ActiveDocument.Bookmarks("tspa_1").Range
        spa_3_7.Text = "S2PA"
        Set SPA_2_7 = ActiveDocument.Bookmarks("tSPA_2").Range
        SPA_2_7.Text = "S2PA"
        Set SPB_7 = ActiveDocument.Bookmarks("tSPB").Range
        SPB_7.Text = "S1PB"
        Set SPB_2_7 = ActiveDocument.Bookmarks("tSPB_2").Range
        SPB_2_7.Text = "S1PB"
        Set SPB_3_7 = ActiveDocument.Bookmarks("tSPB_3").Range
        SPB_3_7.Text = "S1PB"
        Set spc_1_7 = ActiveDocument.Bookmarks("tspc_1").Range
        spc_1_7.Text = "S2PC"
        Set SPC_3_7 = ActiveDocument.Bookmarks("tSPC_3").Range
        SPC_3_7.Text = "S2PC"
        Set SPC_4_7 = ActiveDocument.Bookmarks("tSPC_4").Range
        SPC_4_7.Text = "S2PC"
        
    End If
    
    'If Power System 3 is selected

    If cboPS.value = "Power System 3 " Then
        Set PG_7 = ActiveDocument.Bookmarks("tPG").Range
        PG_7.Text = "PG3"
        Set pg_1_7 = ActiveDocument.Bookmarks("tPG_1").Range
        pg_1_7.Text = "PG3"
        Set PG_2_7 = ActiveDocument.Bookmarks("tPG_2").Range
        PG_2_7.Text = "PG3"
        Set PG_3_7 = ActiveDocument.Bookmarks("tPG_3").Range
        PG_3_7.Text = "PG3"
        Set PG_4_7 = ActiveDocument.Bookmarks("tPG_4").Range
        PG_4_7.Text = "PG3"
        Set PG_5_7 = ActiveDocument.Bookmarks("tPG_5").Range
        PG_5_7.Text = "PG3"
        Set PG_6_7 = ActiveDocument.Bookmarks("tPG_6").Range
        PG_6_7.Text = "PG3"
        Set PG_7_7 = ActiveDocument.Bookmarks("tPG_7").Range
        PG_7_7.Text = "PG3"
        Set USB_2_7 = ActiveDocument.Bookmarks("tUSB").Range
        USB_2_7.Text = "USB3"
        Set USB_3_7 = ActiveDocument.Bookmarks("tUSB_4").Range
        USB_3_7.Text = "USB3"
        Set USB_4_7 = ActiveDocument.Bookmarks("tUSB_1").Range
        USB_4_7.Text = "USB3"
        Set USB_5_7 = ActiveDocument.Bookmarks("tUSB_2").Range
        USB_5_7.Text = "USB3"
        Set PGUSB_7 = ActiveDocument.Bookmarks("tPGUSB").Range
        PGUSB_7.Text = "PG3-USB3"
        Set PGUSB_3_7 = ActiveDocument.Bookmarks("tPGUSB_2").Range
        PGUSB_3_7.Text = "PG3-USB3"
        Set S_7 = ActiveDocument.Bookmarks("tS").Range
        S_7.Text = "S3"
        Set SPA_7 = ActiveDocument.Bookmarks("tSPA").Range
        SPA_7.Text = "S3PA"
        Set spa_3_7 = ActiveDocument.Bookmarks("tspa_1").Range
        spa_3_7.Text = "S3PA"
        Set SPA_2_7 = ActiveDocument.Bookmarks("tSPA_2").Range
        SPA_2_7.Text = "S3PA"
        Set SPB_7 = ActiveDocument.Bookmarks("tSPB").Range
        SPB_7.Text = "S1PB"
        Set SPB_2_7 = ActiveDocument.Bookmarks("tSPB_2").Range
        SPB_2_7.Text = "S3PB"
        Set SPB_3_7 = ActiveDocument.Bookmarks("tSPB_3").Range
        SPB_3_7.Text = "S3PB"
        Set spc_1_7 = ActiveDocument.Bookmarks("tspc_1").Range
        spc_1_7.Text = "S3PC"
        Set SPC_3_7 = ActiveDocument.Bookmarks("tSPC_3").Range
        SPC_3_7.Text = "S3PC"
        Set SPC_4_7 = ActiveDocument.Bookmarks("tSPC_4").Range
        SPC_4_7.Text = "S3PC"
        
    End If
    
    'If Power System 5 is selected

    If cboPS.value = "Power System 5" Then
        Set PG_7 = ActiveDocument.Bookmarks("tPG").Range
        PG_7.Text = "PG5"
        Set pg_1_7 = ActiveDocument.Bookmarks("tPG_1").Range
        pg_1_7.Text = "PG5"
        Set PG_2_7 = ActiveDocument.Bookmarks("tPG_2").Range
        PG_2_7.Text = "PG5"
        Set PG_3_7 = ActiveDocument.Bookmarks("tPG_3").Range
        PG_3_7.Text = "PG5"
        Set PG_4_7 = ActiveDocument.Bookmarks("tPG_4").Range
        PG_4_7.Text = "PG5"
        Set PG_5_7 = ActiveDocument.Bookmarks("tPG_5").Range
        PG_5_7.Text = "PG5"
        Set PG_6_7 = ActiveDocument.Bookmarks("tPG_6").Range
        PG_6_7.Text = "PG5"
        Set PG_7_7 = ActiveDocument.Bookmarks("tPG_7").Range
        PG_7_7.Text = "PG5"
        Set USB_2_7 = ActiveDocument.Bookmarks("tUSB").Range
        USB_2_7.Text = "USB5"
        Set USB_3_7 = ActiveDocument.Bookmarks("tUSB_4").Range
        USB_3_7.Text = "USB5"
        Set USB_4_7 = ActiveDocument.Bookmarks("tUSB_1").Range
        USB_4_7.Text = "USB5"
        Set USB_5_7 = ActiveDocument.Bookmarks("tUSB_2").Range
        USB_5_7.Text = "USB5"
        Set PGUSB_7 = ActiveDocument.Bookmarks("tPGUSB").Range
        PGUSB_7.Text = "PG5-USB5"
        Set PGUSB_3_7 = ActiveDocument.Bookmarks("tPGUSB_2").Range
        PGUSB_3_7.Text = "PG5-USB5"
        Set S_7 = ActiveDocument.Bookmarks("tS").Range
        S_7.Text = "S5"
        Set SPA_7 = ActiveDocument.Bookmarks("tSPA").Range
        SPA_7.Text = "S5PA"
        Set spa_3_7 = ActiveDocument.Bookmarks("tspa_1").Range
        spa_3_7.Text = "S5PA"
        Set SPA_2_7 = ActiveDocument.Bookmarks("tSPA_2").Range
        SPA_2_7.Text = "S5PA"
        Set SPB_7 = ActiveDocument.Bookmarks("tSPB").Range
        SPB_7.Text = "S5PB"
        Set SPB_2_7 = ActiveDocument.Bookmarks("tSPB_2").Range
        SPB_2_7.Text = "S5PB"
        Set SPB_3_7 = ActiveDocument.Bookmarks("tSPB_3").Range
        SPB_3_7.Text = "S5PB"
        Set spc_1_7 = ActiveDocument.Bookmarks("tspc_1").Range
        spc_1_7.Text = "S1PC"
        Set SPC_3_7 = ActiveDocument.Bookmarks("tSPC_3").Range
        SPC_3_7.Text = "S1PC"
        Set SPC_4_7 = ActiveDocument.Bookmarks("tSPC_4").Range
        SPC_4_7.Text = "S5PC"
        
    End If
    
    'If Power System 6 is selected

    If cboPS.value = "Power System 6" Then
        Set PG_7 = ActiveDocument.Bookmarks("tPG").Range
        PG_7.Text = "PG6"
        Set pg_1_7 = ActiveDocument.Bookmarks("tPG_1").Range
        pg_1_7.Text = "PG6"
        Set PG_2_7 = ActiveDocument.Bookmarks("tPG_2").Range
        PG_2_7.Text = "PG6"
        Set PG_3_7 = ActiveDocument.Bookmarks("tPG_3").Range
        PG_3_7.Text = "PG6"
        Set PG_4_7 = ActiveDocument.Bookmarks("tPG_4").Range
        PG_4_7.Text = "PG6"
        Set PG_5_7 = ActiveDocument.Bookmarks("tPG_5").Range
        PG_5_7.Text = "PG6"
        Set PG_6_7 = ActiveDocument.Bookmarks("tPG_6").Range
        PG_6_7.Text = "PG6"
        Set PG_7_7 = ActiveDocument.Bookmarks("tPG_7").Range
        PG_7_7.Text = "PG6"
        Set USB_2_7 = ActiveDocument.Bookmarks("tUSB").Range
        USB_2_7.Text = "USB6"
        Set USB_3_7 = ActiveDocument.Bookmarks("tUSB_4").Range
        USB_3_7.Text = "USB6"
        Set USB_4_7 = ActiveDocument.Bookmarks("tUSB_1").Range
        USB_4_7.Text = "USB6"
        Set USB_5_7 = ActiveDocument.Bookmarks("tUSB_2").Range
        USB_5_7.Text = "USB6"
        Set PGUSB_7 = ActiveDocument.Bookmarks("tPGUSB").Range
        PGUSB_7.Text = "PG6-USB6"
        Set PGUSB_3_7 = ActiveDocument.Bookmarks("tPGUSB_2").Range
        PGUSB_3_7.Text = "PG6-USB6"
        Set S_7 = ActiveDocument.Bookmarks("tS").Range
        S_7.Text = "S6"
        Set SPA_7 = ActiveDocument.Bookmarks("tSPA").Range
        SPA_7.Text = "S6PA"
        Set spa_3_7 = ActiveDocument.Bookmarks("tspa_1").Range
        spa_3_7.Text = "S6PA"
        Set SPA_2_7 = ActiveDocument.Bookmarks("tSPA_2").Range
        SPA_2_7.Text = "S6PA"
        Set SPB_7 = ActiveDocument.Bookmarks("tSPB").Range
        SPB_7.Text = "S6PB"
        Set SPB_2_7 = ActiveDocument.Bookmarks("tSPB_2").Range
        SPB_2_7.Text = "S6PB"
        Set SPB_3_7 = ActiveDocument.Bookmarks("tSPB_3").Range
        SPB_3_7.Text = "S6PB"
        Set spc_1_7 = ActiveDocument.Bookmarks("tspc_1").Range
        spc_1_7.Text = "S6PC"
        Set SPC_3_7 = ActiveDocument.Bookmarks("tSPC_3").Range
        SPC_3_7.Text = "S6PC"
        Set SPC_4_7 = ActiveDocument.Bookmarks("tSPC_4").Range
        SPC_4_7.Text = "S6PC"
        
    End If
    
    'If Power System 7 is selected

    If cboPS.value = "Power System 7" Then
        Set PG_7 = ActiveDocument.Bookmarks("tPG").Range
        PG_7.Text = "PG7"
        Set pg_1_7 = ActiveDocument.Bookmarks("tPG_1").Range
        pg_1_7.Text = "PG7"
        Set PG_2_7 = ActiveDocument.Bookmarks("tPG_2").Range
        PG_2_7.Text = "PG7"
        Set PG_3_7 = ActiveDocument.Bookmarks("tPG_3").Range
        PG_3_7.Text = "PG7"
        Set PG_4_7 = ActiveDocument.Bookmarks("tPG_4").Range
        PG_4_7.Text = "PG7"
        Set PG_5_7 = ActiveDocument.Bookmarks("tPG_5").Range
        PG_5_7.Text = "PG7"
        Set PG_6_7 = ActiveDocument.Bookmarks("tPG_6").Range
        PG_6_7.Text = "PG7"
        Set PG_7_7 = ActiveDocument.Bookmarks("tPG_7").Range
        PG_7_7.Text = "PG7"
        Set USB_2_7 = ActiveDocument.Bookmarks("tUSB").Range
        USB_2_7.Text = "USB7"
        Set USB_3_7 = ActiveDocument.Bookmarks("tUSB_4").Range
        USB_3_7.Text = "USB7"
        Set USB_4_7 = ActiveDocument.Bookmarks("tUSB_1").Range
        USB_4_7.Text = "USB7"
        Set USB_5_7 = ActiveDocument.Bookmarks("tUSB_2").Range
        USB_5_7.Text = "USB7"
        Set PGUSB_7 = ActiveDocument.Bookmarks("tPGUSB").Range
        PGUSB_7.Text = "PG7-USB7"
        Set PGUSB_3_7 = ActiveDocument.Bookmarks("tPGUSB_2").Range
        PGUSB_3_7.Text = "PG7-USB7"
        Set S_7 = ActiveDocument.Bookmarks("tS").Range
        S_7.Text = "S7"
        Set SPA_7 = ActiveDocument.Bookmarks("tSPA").Range
        SPA_7.Text = "S7PA"
        Set spa_3_7 = ActiveDocument.Bookmarks("tspa_1").Range
        spa_3_7.Text = "S7PA"
        Set SPA_2_7 = ActiveDocument.Bookmarks("tSPA_2").Range
        SPA_2_7.Text = "S7PA"
        Set SPB_7 = ActiveDocument.Bookmarks("tSPB").Range
        SPB_7.Text = "S7PB"
        Set SPB_2_7 = ActiveDocument.Bookmarks("tSPB_2").Range
        SPB_2_7.Text = "S7PB"
        Set SPB_3_7 = ActiveDocument.Bookmarks("tSPB_3").Range
        SPB_3_7.Text = "S7PB"
        Set spc_1_7 = ActiveDocument.Bookmarks("tspc_1").Range
        spc_1_7.Text = "S7PC"
        Set SPC_3_7 = ActiveDocument.Bookmarks("tSPC_3").Range
        SPC_3_7.Text = "S7PC"
        Set SPC_4_7 = ActiveDocument.Bookmarks("tSPC_4").Range
        SPC_4_7.Text = "S7PC"
        
    End If
    
    'If Power System 8 is selected

    If cboPS.value = "Power System 8" Then
        Set PG_7 = ActiveDocument.Bookmarks("tPG").Range
        PG_7.Text = "PG8"
        Set pg_1_7 = ActiveDocument.Bookmarks("tPG_1").Range
        pg_1_7.Text = "PG8"
        Set PG_2_7 = ActiveDocument.Bookmarks("tPG_2").Range
        PG_2_7.Text = "PG8"
        Set PG_3_7 = ActiveDocument.Bookmarks("tPG_3").Range
        PG_3_7.Text = "PG8"
        Set PG_4_7 = ActiveDocument.Bookmarks("tPG_4").Range
        PG_4_7.Text = "PG8"
        Set PG_5_7 = ActiveDocument.Bookmarks("tPG_5").Range
        PG_5_7.Text = "PG8"
        Set PG_6_7 = ActiveDocument.Bookmarks("tPG_6").Range
        PG_6_7.Text = "PG8"
        Set PG_7_7 = ActiveDocument.Bookmarks("tPG_7").Range
        PG_7_7.Text = "PG8"
        Set USB_2_7 = ActiveDocument.Bookmarks("tUSB").Range
        USB_2_7.Text = "USB8"
        Set USB_3_7 = ActiveDocument.Bookmarks("tUSB_4").Range
        USB_3_7.Text = "USB8"
        Set USB_4_7 = ActiveDocument.Bookmarks("tUSB_1").Range
        USB_4_7.Text = "USB8"
        Set USB_5_7 = ActiveDocument.Bookmarks("tUSB_2").Range
        USB_5_7.Text = "USB8"
        Set PGUSB_7 = ActiveDocument.Bookmarks("tPGUSB").Range
        PGUSB_7.Text = "PG8-USB8"
        Set PGUSB_3_7 = ActiveDocument.Bookmarks("tPGUSB_2").Range
        PGUSB_3_7.Text = "PG8-USB8"
        Set S_7 = ActiveDocument.Bookmarks("tS").Range
        S_7.Text = "S8"
        Set SPA_7 = ActiveDocument.Bookmarks("tSPA").Range
        SPA_7.Text = "S8PA"
        Set spa_3_7 = ActiveDocument.Bookmarks("tspa_1").Range
        spa_3_7.Text = "S8PA"
        Set SPA_2_7 = ActiveDocument.Bookmarks("tSPA_2").Range
        SPA_2_7.Text = "S8PA"
        Set SPB_7 = ActiveDocument.Bookmarks("tSPB").Range
        SPB_7.Text = "S8PB"
        Set SPB_2_7 = ActiveDocument.Bookmarks("tSPB_2").Range
        SPB_2_7.Text = "S8PB"
        Set SPB_3_7 = ActiveDocument.Bookmarks("tSPB_3").Range
        SPB_3_7.Text = "S8PB"
        Set spc_1_7 = ActiveDocument.Bookmarks("tspc_1").Range
        spc_1_7.Text = "S8PC"
        Set SPC_3_7 = ActiveDocument.Bookmarks("tSPC_3").Range
        SPC_3_7.Text = "S8PC"
        Set SPC_4_7 = ActiveDocument.Bookmarks("tSPC_4").Range
        SPC_4_7.Text = "S8PC"
        
    End If

    Set footerSite_7 = ActiveDocument.Bookmarks("tsite").Range
    footerSite_7.Text = cboSite.value
    
    Set footerSite_7_1 = ActiveDocument.Bookmarks("tsite_1").Range
    footerSite_7_1.Text = cboSite.value
    
    Set projectname_1_7 = ActiveDocument.Bookmarks("tprojectname_1").Range
    projectname_1_7.Text = Me.ProjectName.value
    
    Set projectname_7 = ActiveDocument.Bookmarks("tprojectname").Range
    projectname_7.Text = Me.ProjectName.value
    
    Set Criticalitylevel_7 = ActiveDocument.Bookmarks("tCriticalitylevel").Range
    Criticalitylevel_7.Text = cbocriticalitylevel.value
    
    Dim Projectmanagerphone_7 As Range
    Set Projectmanagerphone_7 = ActiveDocument.Bookmarks("tProjectmanagerphone").Range
    Projectmanagerphone_7.Text = Me.Phonenumber.value
    
    Dim Projectmanagerinitials_7 As Range
    Set Projectmanagerinitials_7 = ActiveDocument.Bookmarks("tProjectManagerinitials").Range
    Projectmanagerinitials_7.Text = Me.Initials.value
    
    Dim ProjectManager_7 As Range
    Set ProjectManager_7 = ActiveDocument.Bookmarks("tProjectManager").Range
    ProjectManager_7.Text = Me.ProjectManager.value
    
    Dim OncallManager_7
    Set OncallManager_7 = ActiveDocument.Bookmarks("tOncallmanager").Range
    OncallManager_7.Text = Me.oncall_1.value
    
    Dim Maintenancewindow_7 As Range
    Set Maintenancewindow_7 = ActiveDocument.Bookmarks("tMaintenancewindow").Range
    Maintenancewindow_7.Text = Me.Maintenancewindow.value
    
    Dim Workorder_2_7 As Range
    Set Workorder_2_7 = ActiveDocument.Bookmarks("tworkorder_1").Range
    Workorder_2_7.Text = Me.Workorder_1.value
    
    Dim Workorder_7
    Set Workorder_7 = ActiveDocument.Bookmarks("tworkorder").Range
    Workorder_7.Text = Me.Workorder_1.value
    
    Dim ticketnumber_7 As Range
    Set ticketnumber_7 = ActiveDocument.Bookmarks("tTicketnumber").Range
    ticketnumber_7.Text = Me.ticketnumber_1.value
    
    Dim completiondate_3_7
    Set completiondate_3_7 = ActiveDocument.Bookmarks("tcompletiondate_1").Range
    completiondate_3_7.Text = Me.completiondate_1.value
    
    Dim complettiondate_2_7 As Range
    Set completiondate_2_7 = ActiveDocument.Bookmarks("tcompletiondate").Range
    completiondate_2_7.Text = Me.completiondate_1.value
    
    Dim startdate_2_7 As Range
    Set startdate_2_7 = ActiveDocument.Bookmarks("tstartdate_1").Range
    startdate_2_7.Text = Me.startdate_1.value
    
    Dim Startdate_7 As Range
    Set Startdate_7 = ActiveDocument.Bookmarks("tStartdate").Range
    Startdate_7.Text = Me.startdate_1.value
    
    Dim Endtime_7 As Range
    Set Endtime_7 = ActiveDocument.Bookmarks("tendtime").Range
    Endtime_7.Text = Me.endtime_1.value
    
    Dim Starttime_7 As Range
    Set Starttime_7 = ActiveDocument.Bookmarks("tstarttime").Range
    Starttime_7.Text = Me.starttime_1.value
    
    SPA_7.Font.AllCaps = True
    spa_3_7.Font.AllCaps = True
    SPA_2_7.Font.AllCaps = True
    SPB_7.Font.AllCaps = True
    SPB_2_7.Font.AllCaps = True
    SPB_3_7.Font.AllCaps = True
    spc_1_7.Font.AllCaps = True
    SPC_3_7.Font.AllCaps = True
    SPC_4_7.Font.AllCaps = True
    USB_2_7.Font.AllCaps = True
End Sub
Public Sub Cancel_Click()

Me.Hide

End Sub
Private Sub generatorselectionLAS7()

'Set Generators

    Dim Gen_1 As Range
    Dim Gen_1_2 As Range
    Dim Gen_1_3 As Range
    Dim Gen_1_6 As Range
    Dim Gen_1_7 As Range
    Dim Gen_2 As Range
    Dim Gen_2_2 As Range
    Dim Gen_2_3 As Range
    Dim Gen_2_6 As Range
    Dim Gen_2_7 As Range
    Dim Gen_3 As Range
    Dim Gen_3_2 As Range
    Dim Gen_3_3 As Range
    Dim Gen_3_6 As Range
    Dim Gen_3_7 As Range
    Dim Gen_4 As Range
    Dim Gen_4_2 As Range
    Dim Gen_4_3 As Range
    Dim Gen_4_6 As Range
    Dim Gen_4_7 As Range
    Dim Gen_5 As Range
    Dim Gen_5_2 As Range
    Dim Gen_5_3 As Range
    Dim Gen_5_6 As Range
    Dim Gen_5_7 As Range
    
If cboSite.value = "LAS 7" Then
    Set Gen_1 = ActiveDocument.Bookmarks("tGen_1").Range
    Set Gen_2 = ActiveDocument.Bookmarks("tGen_2").Range
    Set Gen_3 = ActiveDocument.Bookmarks("tGen_3").Range
    Set Gen_4 = ActiveDocument.Bookmarks("tGen_4").Range
    Set Gen_5 = ActiveDocument.Bookmarks("tGen_5").Range
    Set Gen_1_2 = ActiveDocument.Bookmarks("tGen_1_2").Range
    Set Gen_1_3 = ActiveDocument.Bookmarks("tGen_1_3").Range
    Set Gen_1_6 = ActiveDocument.Bookmarks("tGen_1_6").Range
    Set Gen_1_7 = ActiveDocument.Bookmarks("tGen_1_7").Range
    Set Gen_2_2 = ActiveDocument.Bookmarks("tGen_2_2").Range
    Set Gen_2_3 = ActiveDocument.Bookmarks("tGen_2_3").Range
    Set Gen_2_6 = ActiveDocument.Bookmarks("tGen_2_6").Range
    Set Gen_2_7 = ActiveDocument.Bookmarks("tGen_2_7").Range
    Set Gen_3_2 = ActiveDocument.Bookmarks("tGen_3_2").Range
    Set Gen_3_3 = ActiveDocument.Bookmarks("tGen_3_3").Range
    Set Gen_3_6 = ActiveDocument.Bookmarks("tGen_3_6").Range
    Set Gen_3_7 = ActiveDocument.Bookmarks("tGen_3_7").Range
    Set Gen_4_2 = ActiveDocument.Bookmarks("tGen_4_2").Range
    Set Gen_4_3 = ActiveDocument.Bookmarks("tGen_4_3").Range
    Set Gen_4_6 = ActiveDocument.Bookmarks("tGen_4_6").Range
    Set Gen_4_7 = ActiveDocument.Bookmarks("tGen_4_7").Range
    Set Gen_5_2 = ActiveDocument.Bookmarks("tGen_5_2").Range
    Set Gen_5_3 = ActiveDocument.Bookmarks("tGen_5_3").Range
    Set Gen_5_6 = ActiveDocument.Bookmarks("tGen_5_6").Range
    Set Gen_5_7 = ActiveDocument.Bookmarks("tGen_5_7").Range
End If


    If cboPS.value = "Power System 1 " Then
            Gen_1.Text = "1-1"
            Gen_1_2.Text = "1-1"
            Gen_1_3.Text = "1-1"
            Gen_1_6.Text = "1-1"
            Gen_1_7.Text = "1-1"
            Gen_2.Text = "1-2"
            Gen_2_2.Text = "1-2"
            Gen_2_3.Text = "1-2"
            Gen_2_6.Text = "1-2"
            Gen_2_7.Text = "1-2"
            Gen_3.Text = "1-3"
            Gen_3_2.Text = "1-3"
            Gen_3_3.Text = "1-3"
            Gen_3_6.Text = "1-3"
            Gen_3_7.Text = "1-3"
            Gen_4.Text = "1-4"
            Gen_4_2.Text = "1-4"
            Gen_4_3.Text = "1-4"
            Gen_4_6.Text = "1-4"
            Gen_4_7.Text = "1-4"
            Gen_5.Text = "1-5"
            Gen_5_2.Text = "1-5"
            Gen_5_3.Text = "1-5"
            Gen_5_6.Text = "1-5"
            Gen_5_7.Text = "1-5"
            ElseIf cboPS.value = "Power System 2 " Then
            Gen_1.Text = "2-1"
            Gen_1_2.Text = "2-1"
            Gen_1_3.Text = "2-1"
            Gen_1_6.Text = "2-1"
            Gen_1_7.Text = "2-1"
            Gen_2.Text = "2-2"
            Gen_2_2.Text = "2-2"
            Gen_2_3.Text = "2-2"
            Gen_2_6.Text = "2-2"
            Gen_2_7.Text = "2-2"
            Gen_3.Text = "2-3"
            Gen_3_2.Text = "2-3"
            Gen_3_3.Text = "2-3"
            Gen_3_6.Text = "2-3"
            Gen_3_7.Text = "2-3"
            Gen_4.Text = "2-4"
            Gen_4_2.Text = "2-4"
            Gen_4_3.Text = "2-4"
            Gen_4_6.Text = "2-4"
            Gen_4_7.Text = "2-4"
            Gen_5.Text = "2-5"
            Gen_5_2.Text = "2-5"
            Gen_5_3.Text = "2-5"
            Gen_5_6.Text = "2-5"
            Gen_5_7.Text = "2-5"
            ElseIf cboPS.value = "Power System 3 " Then
            Gen_1.Text = "3-1"
            Gen_1_2.Text = "3-1"
            Gen_1_3.Text = "3-1"
            Gen_1_6.Text = "3-1"
            Gen_1_7.Text = "3-1"
            Gen_2.Text = "3-2"
            Gen_2_2.Text = "3-2"
            Gen_2_3.Text = "3-2"
            Gen_2_6.Text = "3-2"
            Gen_2_7.Text = "3-2"
            Gen_3.Text = "3-3"
            Gen_3_2.Text = "3-3"
            Gen_3_3.Text = "3-3"
            Gen_3_6.Text = "3-3"
            Gen_3_7.Text = "3-3"
            Set Gen_4_2 = Nothing
            Set Gen_4_3 = Nothing
            Set Gen_4_6 = Nothing
            Set Gen_4_7 = Nothing
            Set Gen_5_2 = Nothing
            Set Gen_5_3 = Nothing
            Set Gen_5_6 = Nothing
            Set Gen_5_7 = Nothing
            ElseIf cboPS.value = "Power System 5" Then
            Gen_1.Text = "5-1"
            Gen_1_2.Text = "5-1"
            Gen_1_3.Text = "5-1"
            Gen_1_6.Text = "5-1"
            Gen_1_7.Text = "5-1"
            Gen_2.Text = "5-2"
            Gen_2_2.Text = "5-2"
            Gen_2_3.Text = "5-2"
            Gen_2_6.Text = "5-2"
            Gen_2_7.Text = "5-2"
            Gen_3.Text = "5-3"
            Gen_3_2.Text = "5-3"
            Gen_3_3.Text = "5-3"
            Gen_3_6.Text = "5-3"
            Gen_3_7.Text = "5-3"
            Gen_4.Text = "5-4"
            Gen_4_2.Text = "5-4"
            Gen_4_3.Text = "5-4"
            Gen_4_6.Text = "5-4"
            Gen_4_7.Text = "5-4"
            Gen_5.Text = "5-5"
            Gen_5_2.Text = "5-5"
            Gen_5_3.Text = "5-5"
            Gen_5_6.Text = "5-5"
            Gen_5_7.Text = "5-5"
            ElseIf cboPS.value = "Power System 6" Then
            Gen_1.Text = "6-1"
            Gen_1_2.Text = "6-1"
            Gen_1_3.Text = "6-1"
            Gen_1_6.Text = "6-1"
            Gen_1_7.Text = "6-1"
            Gen_2.Text = "6-2"
            Gen_2_2.Text = "6-2"
            Gen_2_3.Text = "6-2"
            Gen_2_6.Text = "6-2"
            Gen_2_7.Text = "6-2"
            Gen_3.Text = "6-3"
            Gen_3_2.Text = "6-3"
            Gen_3_3.Text = "6-3"
            Gen_3_6.Text = "6-3"
            Gen_3_7.Text = "6-3"
            Gen_4.Text = "6-4"
            Gen_4_2.Text = "6-4"
            Gen_4_3.Text = "6-4"
            Gen_4_6.Text = "6-4"
            Gen_4_7.Text = "6-4"
            Gen_5.Text = "6-5"
            Gen_5_2.Text = "6-5"
            Gen_5_3.Text = "6-5"
            Gen_5_6.Text = "6-5"
            Gen_5_7.Text = "6-5"
            ElseIf cboPS.value = "Power System 7" Then
            Gen_1.Text = "7-1"
            Gen_1_2.Text = "7-1"
            Gen_1_3.Text = "7-1"
            Gen_1_6.Text = "7-1"
            Gen_1_7.Text = "7-1"
            Gen_2.Text = "7-2"
            Gen_2_2.Text = "7-2"
            Gen_2_3.Text = "7-2"
            Gen_2_6.Text = "7-2"
            Gen_2_7.Text = "7-2"
            Gen_3.Text = "7-3"
            Gen_3_2.Text = "7-3"
            Gen_3_3.Text = "7-3"
            Gen_3_6.Text = "7-3"
            Gen_3_7.Text = "7-3"
            Set Gen_4_2 = Nothing
            Set Gen_4_3 = Nothing
            Set Gen_4_6 = Nothing
            Set Gen_4_7 = Nothing
            Set Gen_5_2 = Nothing
            Set Gen_5_3 = Nothing
            Set Gen_5_6 = Nothing
            Set Gen_5_7 = Nothing
            ElseIf cboPS.value = "Power System 8" Then
            Gen_1.Text = "8-1"
            Gen_1_2.Text = "8-1"
            Gen_1_3.Text = "8-1"
            Gen_1_6.Text = "8-1"
            Gen_1_7.Text = "8-1"
            Gen_2.Text = "8-2"
            Gen_2_2.Text = "8-2"
            Gen_2_3.Text = "8-2"
            Gen_2_6.Text = "8-2"
            Gen_2_7.Text = "8-2"
            Gen_3.Text = "8-3"
            Gen_3_2.Text = "8-3"
            Gen_3_3.Text = "8-3"
            Gen_3_6.Text = "8-3"
            Gen_3_7.Text = "8-3"
            Gen_4.Text = "8-4"
            Gen_4_2.Text = "8-4"
            Gen_4_3.Text = "8-4"
            Gen_4_6.Text = "8-4"
            Gen_4_7.Text = "8-4"
            Gen_5.Text = "8-5"
            Gen_5_2.Text = "8-5"
            Gen_5_3.Text = "8-5"
            Gen_5_6.Text = "8-5"
            Gen_5_7.Text = "8-5"
    End If

End Sub
Private Sub generatorSelection()

'Set Generators when a load bank will be involved with the maintenance.

    Dim Gen_1 As Range
    Dim Gen_1_2 As Range
    Dim Gen_1_3 As Range
    Dim Gen_1_6 As Range
    Dim Gen_1_7 As Range
    Dim Gen_2 As Range
    Dim Gen_2_2 As Range
    Dim Gen_2_3 As Range
    Dim Gen_2_6 As Range
    Dim Gen_2_7 As Range
    Dim Gen_3 As Range
    Dim Gen_3_2 As Range
    Dim Gen_3_3 As Range
    Dim Gen_3_6 As Range
    Dim Gen_3_7 As Range
    Dim Gen_4 As Range
    Dim Gen_4_2 As Range
    Dim Gen_4_3 As Range
    Dim Gen_4_6 As Range
    Dim Gen_4_7 As Range
    Dim Gen_5 As Range
    Dim Gen_5_2 As Range
    Dim Gen_5_3 As Range
    Dim Gen_5_6 As Range
    Dim Gen_5_7 As Range
    Dim Gen_6 As Range
    Dim Gen_6_2 As Range
    Dim Gen_6_3 As Range
    Dim Gen_6_6 As Range
    Dim Gen_6_7 As Range
    
If cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11" Then
    Set Gen_1 = ActiveDocument.Bookmarks("tGen_1").Range
    Set Gen_2 = ActiveDocument.Bookmarks("tGen_2").Range
    Set Gen_3 = ActiveDocument.Bookmarks("tGen_3").Range
    Set Gen_4 = ActiveDocument.Bookmarks("tGen_4").Range
    Set Gen_5 = ActiveDocument.Bookmarks("tGen_5").Range
    Set Gen_6 = ActiveDocument.Bookmarks("tGen_6").Range
    Set Gen_1_2 = ActiveDocument.Bookmarks("tGen_1_2").Range
    Set Gen_1_3 = ActiveDocument.Bookmarks("tGen_1_3").Range
    Set Gen_1_6 = ActiveDocument.Bookmarks("tGen_1_6").Range
    Set Gen_1_7 = ActiveDocument.Bookmarks("tGen_1_7").Range
    Set Gen_2_2 = ActiveDocument.Bookmarks("tGen_2_2").Range
    Set Gen_2_3 = ActiveDocument.Bookmarks("tGen_2_3").Range
    Set Gen_2_6 = ActiveDocument.Bookmarks("tGen_2_6").Range
    Set Gen_2_7 = ActiveDocument.Bookmarks("tGen_2_7").Range
    Set Gen_3_2 = ActiveDocument.Bookmarks("tGen_3_2").Range
    Set Gen_3_3 = ActiveDocument.Bookmarks("tGen_3_3").Range
    Set Gen_3_6 = ActiveDocument.Bookmarks("tGen_3_6").Range
    Set Gen_3_7 = ActiveDocument.Bookmarks("tGen_3_7").Range
    Set Gen_4_2 = ActiveDocument.Bookmarks("tGen_4_2").Range
    Set Gen_4_3 = ActiveDocument.Bookmarks("tGen_4_3").Range
    Set Gen_4_6 = ActiveDocument.Bookmarks("tGen_4_6").Range
    Set Gen_4_7 = ActiveDocument.Bookmarks("tGen_4_7").Range
    Set Gen_5_2 = ActiveDocument.Bookmarks("tGen_5_2").Range
    Set Gen_5_3 = ActiveDocument.Bookmarks("tGen_5_3").Range
    Set Gen_5_6 = ActiveDocument.Bookmarks("tGen_5_6").Range
    Set Gen_5_7 = ActiveDocument.Bookmarks("tGen_5_7").Range
    Set Gen_6_2 = ActiveDocument.Bookmarks("tGen_6_2").Range
    Set Gen_6_3 = ActiveDocument.Bookmarks("tGen_6_3").Range
    Set Gen_6_6 = ActiveDocument.Bookmarks("tGen_6_6").Range
    Set Gen_6_7 = ActiveDocument.Bookmarks("tGen_6_7").Range
    End If

    If cboPS.value = "Power System 11" Then
            Gen_1.Text = "11-1"
            Gen_1_2.Text = "11-1"
            Gen_1_3.Text = "11-1"
            Gen_1_6.Text = "11-1"
            Gen_1_7.Text = "11-1"
            Gen_2.Text = "11-2"
            Gen_2_2.Text = "11-2"
            Gen_2_3.Text = "11-2"
            Gen_2_6.Text = "11-2"
            Gen_2_7.Text = "11-2"
            Gen_3.Text = "11-3"
            Gen_3_2.Text = "11-3"
            Gen_3_3.Text = "11-3"
            Gen_3_6.Text = "11-3"
            Gen_3_7.Text = "11-3"
            Gen_4.Text = "11-4"
            Gen_4_2.Text = "11-4"
            Gen_4_3.Text = "11-4"
            Gen_4_6.Text = "11-4"
            Gen_4_7.Text = "11-4"
            Gen_5.Text = "11-5"
            Gen_5_2.Text = "11-5"
            Gen_5_3.Text = "11-5"
            Gen_5_6.Text = "11-5"
            Gen_5_7.Text = "11-5"
            Gen_6.Text = "11-6"
            Gen_6_2.Text = "11-6"
            Gen_6_3.Text = "11-6"
            Gen_6_6.Text = "11-6"
            Gen_6_7.Text = "11-6"
        ElseIf cboPS.value = "Power System 12" Then
            Gen_1.Text = "12-1"
            Gen_1_2.Text = "12-1"
            Gen_1_3.Text = "12-1"
            Gen_1_6.Text = "12-1"
            Gen_1_7.Text = "12-1"
            Gen_2.Text = "12-2"
            Gen_2_2.Text = "12-2"
            Gen_2_3.Text = "12-2"
            Gen_2_6.Text = "12-2"
            Gen_2_7.Text = "12-2"
            Gen_3.Text = "12-3"
            Gen_3_2.Text = "12-3"
            Gen_3_3.Text = "12-3"
            Gen_3_6.Text = "12-3"
            Gen_3_7.Text = "12-3"
            Gen_4.Text = "12-4"
            Gen_4_2.Text = "12-4"
            Gen_4_3.Text = "12-4"
            Gen_4_6.Text = "12-4"
            Gen_4_7.Text = "12-4"
            Gen_5.Text = "12-5"
            Gen_5_2.Text = "12-5"
            Gen_5_3.Text = "12-5"
            Gen_5_6.Text = "12-5"
            Gen_5_7.Text = "12-5"
            Gen_6.Text = "12-6"
            Gen_6_2.Text = "12-6"
            Gen_6_3.Text = "12-6"
            Gen_6_6.Text = "12-6"
            Gen_6_7.Text = "12-6"
        ElseIf cboPS.value = "Power System 13" Then
            Gen_1.Text = "13-1"
            Gen_1_2.Text = "13-1"
            Gen_1_3.Text = "13-1"
            Gen_1_6.Text = "13-1"
            Gen_1_7.Text = "13-1"
            Gen_2.Text = "13-2"
            Gen_2_2.Text = "13-2"
            Gen_2_3.Text = "13-2"
            Gen_2_6.Text = "13-2"
            Gen_2_7.Text = "13-2"
            Gen_3.Text = "13-3"
            Gen_3_2.Text = "13-3"
            Gen_3_3.Text = "13-3"
            Gen_3_6.Text = "13-3"
            Gen_3_7.Text = "13-3"
            Gen_4.Text = "13-4"
            Gen_4_2.Text = "13-4"
            Gen_4_3.Text = "13-4"
            Gen_4_6.Text = "13-4"
            Gen_4_7.Text = "13-4"
            Gen_5.Text = "13-5"
            Gen_5_2.Text = "13-5"
            Gen_5_3.Text = "13-5"
            Gen_5_6.Text = "13-5"
            Gen_5_7.Text = "13-5"
            Gen_6.Text = "13-6"
            Gen_6_2.Text = "13-6"
            Gen_6_3.Text = "13-6"
            Gen_6_6.Text = "13-6"
            Gen_6_7.Text = "13-6"
        ElseIf cboPS.value = "Power System 14" Then
            Gen_1.Text = "14-1"
            Gen_1_2.Text = "14-1"
            Gen_1_3.Text = "14-1"
            Gen_1_6.Text = "14-1"
            Gen_1_7.Text = "14-1"
            Gen_2.Text = "14-2"
            Gen_2_2.Text = "14-2"
            Gen_2_3.Text = "14-2"
            Gen_2_6.Text = "14-2"
            Gen_2_7.Text = "14-2"
            Gen_3.Text = "14-3"
            Gen_3_2.Text = "14-3"
            Gen_3_3.Text = "14-3"
            Gen_3_6.Text = "14-3"
            Gen_3_7.Text = "14-3"
            Gen_4.Text = "14-4"
            Gen_4_2.Text = "14-4"
            Gen_4_3.Text = "14-4"
            Gen_4_6.Text = "14-4"
            Gen_4_7.Text = "14-4"
            Gen_5.Text = "14-5"
            Gen_5_2.Text = "14-5"
            Gen_5_3.Text = "14-5"
            Gen_5_6.Text = "14-5"
            Gen_5_7.Text = "14-5"
            Gen_6.Text = "14-6"
            Gen_6_2.Text = "14-6"
            Gen_6_3.Text = "14-6"
            Gen_6_6.Text = "14-6"
            Gen_6_7.Text = "14-6"
        ElseIf cboPS.value = "Power System 1  " Then
            Gen_1.Text = "1-1"
            Gen_1_2.Text = "1-1"
            Gen_1_3.Text = "1-1"
            Gen_1_6.Text = "1-1"
            Gen_1_7.Text = "1-1"
            Gen_2.Text = "1-2"
            Gen_2_2.Text = "1-2"
            Gen_2_3.Text = "1-2"
            Gen_2_6.Text = "1-2"
            Gen_2_7.Text = "1-2"
            Gen_3.Text = "1-3"
            Gen_3_2.Text = "1-3"
            Gen_3_3.Text = "1-3"
            Gen_3_6.Text = "1-3"
            Gen_3_7.Text = "1-3"
            Gen_4.Text = "1-4"
            Gen_4_2.Text = "1-4"
            Gen_4_3.Text = "1-4"
            Gen_4_6.Text = "1-4"
            Gen_4_7.Text = "1-4"
            Gen_5.Text = "1-5"
            Gen_5_2.Text = "1-5"
            Gen_5_3.Text = "1-5"
            Gen_5_6.Text = "1-5"
            Gen_5_7.Text = "1-5"
            Gen_6.Text = "1-6"
            Gen_6_2.Text = "1-6"
            Gen_6_3.Text = "1-6"
            Gen_6_6.Text = "1-6"
            Gen_6_7.Text = "1-6"
        ElseIf cboPS.value = "Power System 2  " Then
            Gen_1.Text = "2-1"
            Gen_1_2.Text = "2-1"
            Gen_1_3.Text = "2-1"
            Gen_1_6.Text = "2-1"
            Gen_1_7.Text = "2-1"
            Gen_2.Text = "2-2"
            Gen_2_2.Text = "2-2"
            Gen_2_3.Text = "2-2"
            Gen_2_6.Text = "2-2"
            Gen_2_7.Text = "2-2"
            Gen_3.Text = "2-3"
            Gen_3_2.Text = "2-3"
            Gen_3_3.Text = "2-3"
            Gen_3_6.Text = "2-3"
            Gen_3_7.Text = "2-3"
            Gen_4.Text = "2-4"
            Gen_4_2.Text = "2-4"
            Gen_4_3.Text = "2-4"
            Gen_4_6.Text = "2-4"
            Gen_4_7.Text = "2-4"
            Gen_5.Text = "2-5"
            Gen_5_2.Text = "2-5"
            Gen_5_3.Text = "2-5"
            Gen_5_6.Text = "2-5"
            Gen_5_7.Text = "2-5"
            Gen_6.Text = "2-6"
            Gen_6_2.Text = "2-6"
            Gen_6_3.Text = "2-6"
            Gen_6_6.Text = "2-6"
            Gen_6_7.Text = "2-6"
        ElseIf cboPS.value = "Power System 3  " Then
            Gen_1.Text = "3-1"
            Gen_1_2.Text = "3-1"
            Gen_1_3.Text = "3-1"
            Gen_1_6.Text = "3-1"
            Gen_1_7.Text = "3-1"
            Gen_2.Text = "3-2"
            Gen_2_2.Text = "3-2"
            Gen_2_3.Text = "3-2"
            Gen_2_6.Text = "3-2"
            Gen_2_7.Text = "3-2"
            Gen_3.Text = "3-3"
            Gen_3_2.Text = "3-3"
            Gen_3_3.Text = "3-3"
            Gen_3_6.Text = "3-3"
            Gen_3_7.Text = "3-3"
            Gen_4.Text = "3-4"
            Gen_4_2.Text = "3-4"
            Gen_4_3.Text = "3-4"
            Gen_4_6.Text = "3-4"
            Gen_4_7.Text = "3-4"
            Gen_5.Text = "3-5"
            Gen_5_2.Text = "3-5"
            Gen_5_3.Text = "3-5"
            Gen_5_6.Text = "3-5"
            Gen_5_7.Text = "3-5"
            Gen_6.Text = "3-6"
            Gen_6_2.Text = "3-6"
            Gen_6_3.Text = "3-6"
            Gen_6_6.Text = "3-6"
            Gen_6_7.Text = "3-6"
        ElseIf cboPS.value = "Power System 4  " Then
            Gen_1.Text = "4-1"
            Gen_1_2.Text = "4-1"
            Gen_1_3.Text = "4-1"
            Gen_1_6.Text = "4-1"
            Gen_1_7.Text = "4-1"
            Gen_2.Text = "4-2"
            Gen_2_2.Text = "4-2"
            Gen_2_3.Text = "4-2"
            Gen_2_6.Text = "4-2"
            Gen_2_7.Text = "4-2"
            Gen_3.Text = "4-3"
            Gen_3_2.Text = "4-3"
            Gen_3_3.Text = "4-3"
            Gen_3_6.Text = "4-3"
            Gen_3_7.Text = "4-3"
            Gen_4.Text = "4-4"
            Gen_4_2.Text = "4-4"
            Gen_4_3.Text = "4-4"
            Gen_4_6.Text = "4-4"
            Gen_4_7.Text = "4-4"
            Gen_5.Text = "4-5"
            Gen_5_2.Text = "4-5"
            Gen_5_3.Text = "4-5"
            Gen_5_6.Text = "4-5"
            Gen_5_7.Text = "4-5"
            Gen_6.Text = "4-6"
            Gen_6_2.Text = "4-6"
            Gen_6_3.Text = "4-6"
            Gen_6_6.Text = "4-6"
            Gen_6_7.Text = "4-6"
        ElseIf cboPS.value = " Power System 1" Then
            Gen_1.Text = "1-1"
            Gen_1_2.Text = "1-1"
            Gen_1_3.Text = "1-1"
            Gen_1_6.Text = "1-1"
            Gen_1_7.Text = "1-1"
            Gen_2.Text = "1-2"
            Gen_2_2.Text = "1-2"
            Gen_2_3.Text = "1-2"
            Gen_2_6.Text = "1-2"
            Gen_2_7.Text = "1-2"
            Gen_3.Text = "1-3"
            Gen_3_2.Text = "1-3"
            Gen_3_3.Text = "1-3"
            Gen_3_6.Text = "1-3"
            Gen_3_7.Text = "1-3"
            Gen_4.Text = "1-4"
            Gen_4_2.Text = "1-4"
            Gen_4_3.Text = "1-4"
            Gen_4_6.Text = "1-4"
            Gen_4_7.Text = "1-4"
            Gen_5.Text = "1-5"
            Gen_5_2.Text = "1-5"
            Gen_5_3.Text = "1-5"
            Gen_5_6.Text = "1-5"
            Gen_5_7.Text = "1-5"
            Gen_6.Text = "1-6"
            Gen_6_2.Text = "1-6"
            Gen_6_3.Text = "1-6"
            Gen_6_6.Text = "1-6"
            Gen_6_7.Text = "1-6"
        ElseIf cboPS.value = " Power System 2" Then
            Gen_1.Text = "2-1"
            Gen_1_2.Text = "2-1"
            Gen_1_3.Text = "2-1"
            Gen_1_6.Text = "2-1"
            Gen_1_7.Text = "2-1"
            Gen_2.Text = "2-2"
            Gen_2_2.Text = "2-2"
            Gen_2_3.Text = "2-2"
            Gen_2_6.Text = "2-2"
            Gen_2_7.Text = "2-2"
            Gen_3.Text = "2-3"
            Gen_3_2.Text = "2-3"
            Gen_3_3.Text = "2-3"
            Gen_3_6.Text = "2-3"
            Gen_3_7.Text = "2-3"
            Gen_4.Text = "2-4"
            Gen_4_2.Text = "2-4"
            Gen_4_3.Text = "2-4"
            Gen_4_6.Text = "2-4"
            Gen_4_7.Text = "2-4"
            Gen_5.Text = "2-5"
            Gen_5_2.Text = "2-5"
            Gen_5_3.Text = "2-5"
            Gen_5_6.Text = "2-5"
            Gen_5_7.Text = "2-5"
            Gen_6.Text = "2-6"
            Gen_6_2.Text = "2-6"
            Gen_6_3.Text = "2-6"
            Gen_6_6.Text = "2-6"
            Gen_6_7.Text = "2-6"
        ElseIf cboPS.value = " Power System 3" Then
            Gen_1.Text = "3-1"
            Gen_1_2.Text = "3-1"
            Gen_1_3.Text = "3-1"
            Gen_1_6.Text = "3-1"
            Gen_1_7.Text = "3-1"
            Gen_2.Text = "3-2"
            Gen_2_2.Text = "3-2"
            Gen_2_3.Text = "3-2"
            Gen_2_6.Text = "3-2"
            Gen_2_7.Text = "3-2"
            Gen_3.Text = "3-3"
            Gen_3_2.Text = "3-3"
            Gen_3_3.Text = "3-3"
            Gen_3_6.Text = "3-3"
            Gen_3_7.Text = "3-3"
            Gen_4.Text = "11-4"
            Gen_4_2.Text = "3-4"
            Gen_4_3.Text = "3-4"
            Gen_4_6.Text = "3-4"
            Gen_4_7.Text = "3-4"
            Gen_5.Text = "3-5"
            Gen_5_2.Text = "3-5"
            Gen_5_3.Text = "3-5"
            Gen_5_6.Text = "3-5"
            Gen_5_7.Text = "3-5"
            Gen_6.Text = "3-6"
            Gen_6_2.Text = "3-6"
            Gen_6_3.Text = "3-6"
            Gen_6_6.Text = "3-6"
            Gen_6_7.Text = "3-6"
        ElseIf cboPS.value = "Power System 1" Then
            Gen_1.Text = "1-1"
            Gen_1_2.Text = "1-1"
            Gen_1_3.Text = "1-1"
            Gen_1_6.Text = "1-1"
            Gen_1_7.Text = "1-1"
            Gen_2.Text = "1-2"
            Gen_2_2.Text = "1-2"
            Gen_2_3.Text = "1-2"
            Gen_2_6.Text = "1-2"
            Gen_2_7.Text = "1-2"
            Gen_3.Text = "1-3"
            Gen_3_2.Text = "1-3"
            Gen_3_3.Text = "1-3"
            Gen_3_6.Text = "1-3"
            Gen_3_7.Text = "1-3"
            Gen_4.Text = "1-4"
            Gen_4_2.Text = "1-4"
            Gen_4_3.Text = "1-4"
            Gen_4_6.Text = "1-4"
            Gen_4_7.Text = "1-4"
            Gen_5.Text = "1-5"
            Gen_5_2.Text = "1-5"
            Gen_5_3.Text = "1-5"
            Gen_5_6.Text = "1-5"
            Gen_5_7.Text = "1-5"
            Gen_6.Text = "1-6"
            Gen_6_2.Text = "1-6"
            Gen_6_3.Text = "1-6"
            Gen_6_6.Text = "1-6"
            Gen_6_7.Text = "1-6"
        ElseIf cboPS.value = "Power System 2" Then
            Gen_1.Text = "2-1"
            Gen_1_2.Text = "2-1"
            Gen_1_3.Text = "2-1"
            Gen_1_6.Text = "2-1"
            Gen_1_7.Text = "2-1"
            Gen_2.Text = "2-2"
            Gen_2_2.Text = "2-2"
            Gen_2_3.Text = "2-2"
            Gen_2_6.Text = "2-2"
            Gen_2_7.Text = "2-2"
            Gen_3.Text = "2-3"
            Gen_3_2.Text = "2-3"
            Gen_3_3.Text = "2-3"
            Gen_3_6.Text = "2-3"
            Gen_3_7.Text = "2-3"
            Gen_4.Text = "2-4"
            Gen_4_2.Text = "2-4"
            Gen_4_3.Text = "2-4"
            Gen_4_6.Text = "2-4"
            Gen_4_7.Text = "2-4"
            Gen_5.Text = "2-5"
            Gen_5_2.Text = "2-5"
            Gen_5_3.Text = "2-5"
            Gen_5_6.Text = "2-5"
            Gen_5_7.Text = "2-5"
            Gen_6.Text = "2-6"
            Gen_6_2.Text = "2-6"
            Gen_6_3.Text = "2-6"
            Gen_6_6.Text = "2-6"
            Gen_6_7.Text = "2-6"
        ElseIf cboPS.value = "Power System 3" Then
            Gen_1.Text = "3-1"
            Gen_1_2.Text = "3-1"
            Gen_1_3.Text = "3-1"
            Gen_1_6.Text = "3-1"
            Gen_1_7.Text = "3-1"
            Gen_2.Text = "3-2"
            Gen_2_2.Text = "3-2"
            Gen_2_3.Text = "3-2"
            Gen_2_6.Text = "3-2"
            Gen_2_7.Text = "3-2"
            Gen_3.Text = "3-3"
            Gen_3_2.Text = "3-3"
            Gen_3_3.Text = "3-3"
            Gen_3_6.Text = "3-3"
            Gen_3_7.Text = "3-3"
            Gen_4.Text = "3-4"
            Gen_4_2.Text = "3-4"
            Gen_4_3.Text = "3-4"
            Gen_4_6.Text = "3-4"
            Gen_4_7.Text = "3-4"
            Gen_5.Text = "3-5"
            Gen_5_2.Text = "3-5"
            Gen_5_3.Text = "3-5"
            Gen_5_6.Text = "3-5"
            Gen_5_7.Text = "3-5"
            Gen_6.Text = "3-6"
            Gen_6_2.Text = "3-6"
            Gen_6_3.Text = "3-6"
            Gen_6_6.Text = "3-6"
            Gen_6_7.Text = "3-6"
    End If

   If cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal" _
    Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal" Then
        Set BS1 = ActiveDocument.Bookmarks("tBS1").Range
        Set BS1_1 = ActiveDocument.Bookmarks("tBS1_1").Range
        Set GSBA = ActiveDocument.Bookmarks("tGSBA").Range
        Set GSBA_1 = ActiveDocument.Bookmarks("tGSBA_1").Range
        Set GSBA_2 = ActiveDocument.Bookmarks("tGSBA_2").Range
        Set GSBA_3 = ActiveDocument.Bookmarks("tGSBA_3").Range
        Set GSBA_4 = ActiveDocument.Bookmarks("tGSBA_4").Range
        Set GSBA_5 = ActiveDocument.Bookmarks("tGSBA_5").Range
        Set GSBB = ActiveDocument.Bookmarks("tGSBB").Range
        Set GSBB_1 = ActiveDocument.Bookmarks("tGSBB_1").Range
        Set GSBB_2 = ActiveDocument.Bookmarks("tGSBB_2").Range
        Set GSBB_3 = ActiveDocument.Bookmarks("tGSBB_3").Range
        Set GSBB_4 = ActiveDocument.Bookmarks("tGSBB_4").Range
        Set GSBB_5 = ActiveDocument.Bookmarks("tGSBB_5").Range
        Set GSBB_6 = ActiveDocument.Bookmarks("tGSBB_6").Range
        Set GSBC = ActiveDocument.Bookmarks("tGSBC").Range
        Set GSBC_1 = ActiveDocument.Bookmarks("tGSBC_1").Range
        Set GSBC_2 = ActiveDocument.Bookmarks("tGSBC_2").Range
        Set GSBC_3 = ActiveDocument.Bookmarks("tGSBC_3").Range
        Set GSBC_4 = ActiveDocument.Bookmarks("tGSBC_4").Range
        Set GSBC_5 = ActiveDocument.Bookmarks("tGSBC_5").Range
        Set SPA = ActiveDocument.Bookmarks("tSPA").Range
        Set spa_3 = ActiveDocument.Bookmarks("tspa_1").Range
        Set SPA_2 = ActiveDocument.Bookmarks("tSPA_2").Range
        Set SPB = ActiveDocument.Bookmarks("tSPB").Range
        Set SPB_2 = ActiveDocument.Bookmarks("tSPB_2").Range
        Set SPB_3 = ActiveDocument.Bookmarks("tSPB_3").Range
        Set spc_1 = ActiveDocument.Bookmarks("tspc_1").Range
        Set SPC_3 = ActiveDocument.Bookmarks("tSPC_3").Range
        Set SPC_4 = ActiveDocument.Bookmarks("tSPC_4").Range
    End If

'If Power System 11 is selected
    If cboPS.value = "Power System 11" Then
        BS1.Text = "11S-BS1"
        BS1_1.Text = "11S-BS1"
        GSBA.Text = "11GSBA"
        GSBA_1.Text = "11GSBA"
        GSBA_2.Text = "11GSBA"
        GSBA_3.Text = "11GSBA"
        GSBA_4.Text = "11GSBA"
        GSBA_5.Text = "11GSBA"
        GSBB.Text = "11GSBB"
        GSBB_1.Text = "11GSBB"
        GSBB_2.Text = "11GSBB"
        GSBB_3.Text = "11GSBB"
        GSBB_4.Text = "11GSBB"
        GSBB_5.Text = "11GSBB"
        GSBB_6.Text = "11GSBB"
        GSBC.Text = "11GSBC"
        GSBC_1.Text = "11GSBC"
        GSBC_2.Text = "11GSBC"
        GSBC_3.Text = "11GSBC"
        GSBC_4.Text = "11GSBC"
        GSBC_5.Text = "11GSBC"
        SPA.Text = "11MVSA"
        spa_3.Text = "11MVSA"
        SPA_2.Text = "11MVSA"
        SPB.Text = "11MVSB"
        SPB_2.Text = "11MVSB"
        SPB_3.Text = "11MVSB"
        spc_1.Text = "11MVSC"
        SPC_3.Text = "11MVSC"
        SPC_4.Text = "11MVSC"
    End If

'If Power System 12 is selected
    If cboPS.value = "Power System 12" Then
        BS1.Text = "12S-BS1"
        BS1_1.Text = "12S-BS1"
        GSBA.Text = "12GSBA"
        GSBA_1.Text = "12GSBA"
        GSBA_2.Text = "12GSBA"
        GSBA_3.Text = "12GSBA"
        GSBA_4.Text = "12GSBA"
        GSBA_5.Text = "12GSBA"
        GSBB.Text = "12GSBB"
        GSBB_1.Text = "12GSBB"
        GSBB_2.Text = "12GSBB"
        GSBB_3.Text = "12GSBB"
        GSBB_4.Text = "12GSBB"
        GSBB_5.Text = "12GSBB"
        GSBB_6.Text = "12GSBB"
        GSBC.Text = "12GSBC"
        GSBC_1.Text = "12GSBC"
        GSBC_2.Text = "12GSBC"
        GSBC_3.Text = "12GSBC"
        GSBC_4.Text = "12GSBC"
        GSBC_5.Text = "12GSBC"
        SPA.Text = "12MVSA"
        spa_3.Text = "12MVSA"
        SPA_2.Text = "12MVSA"
        SPB.Text = "12MVSB"
        SPB_2.Text = "12MVSB"
        SPB_3.Text = "12MVSB"
        spc_1.Text = "12MVSC"
        SPC_3.Text = "12MVSC"
        SPC_4.Text = "12MVSC"
    End If
'If Power System 13 is selected
    If cboPS.value = "Power System 13" Then
        BS1.Text = "13S-BS1"
        BS1_1.Text = "13S-BS1"
        GSBA.Text = "13GSBA"
        GSBA_1.Text = "13GSBA"
        GSBA_2.Text = "13GSBA"
        GSBA_3.Text = "13GSBA"
        GSBA_4.Text = "13GSBA"
        GSBA_5.Text = "13GSBA"
        GSBB.Text = "13GSBB"
        GSBB_1.Text = "13GSBB"
        GSBB_2.Text = "13GSBB"
        GSBB_3.Text = "13GSBB"
        GSBB_4.Text = "13GSBB"
        GSBB_5.Text = "13GSBB"
        GSBB_6.Text = "13GSBB"
        GSBC.Text = "13GSBC"
        GSBC_1.Text = "13GSBC"
        GSBC_2.Text = "13GSBC"
        GSBC_3.Text = "13GSBC"
        GSBC_4.Text = "13GSBC"
        GSBC_5.Text = "13GSBC"
        SPA.Text = "13MVSA"
        spa_3.Text = "13MVSA"
        SPA_2.Text = "13MVSA"
        SPB.Text = "13MVSB"
        SPB_2.Text = "13MVSB"
        SPB_3.Text = "13MVSB"
        spc_1.Text = "13MVSC"
        SPC_3.Text = "13MVSC"
        SPC_4.Text = "13MVSC"
    End If
'If Power System 14 is selected
    If cboPS.value = "Power System 14" Then
        BS1.Text = "14S-BS1"
        BS1_1.Text = "14S-BS1"
        GSBA.Text = "14GSBA"
        GSBA_1.Text = "14GSBA"
        GSBA_2.Text = "14GSBA"
        GSBA_3.Text = "14GSBA"
        GSBA_4.Text = "14GSBA"
        GSBA_5.Text = "14GSBA"
        GSBB.Text = "14GSBB"
        GSBB_1.Text = "14GSBB"
        GSBB_2.Text = "14GSBB"
        GSBB_3.Text = "14GSBB"
        GSBB_4.Text = "14GSBB"
        GSBB_5.Text = "14GSBB"
        GSBB_6.Text = "14GSBB"
        GSBC.Text = "14GSBC"
        GSBC_1.Text = "14GSBC"
        GSBC_2.Text = "14GSBC"
        GSBC_3.Text = "14GSBC"
        GSBC_4.Text = "14GSBC"
        GSBC_5.Text = "14GSBC"
        SPA.Text = "14MVSA"
        spa_3.Text = "14MVSA"
        SPA_2.Text = "14MVSA"
        SPB.Text = "14MVSB"
        SPB_2.Text = "14MVSB"
        SPB_3.Text = "14MVSB"
        spc_1.Text = "14MVSC"
        SPC_3.Text = "14MVSC"
        SPC_4.Text = "14MVSC"
    End If
'If LAS 9 Power System 1 is selected
    If cboPS.value = "Power System 1  " Then
        BS1.Text = "15S-BS1"
        BS1_1.Text = "15S-BS1"
        GSBA.Text = "15GSBA"
        GSBA_1.Text = "15GSBA"
        GSBA_2.Text = "15GSBA"
        GSBA_3.Text = "15GSBA"
        GSBA_4.Text = "15GSBA"
        GSBA_5.Text = "15GSBA"
        GSBB.Text = "15GSBB"
        GSBB_1.Text = "15GSBB"
        GSBB_2.Text = "15GSBB"
        GSBB_3.Text = "15GSBB"
        GSBB_4.Text = "15GSBB"
        GSBB_5.Text = "15GSBB"
        GSBB_6.Text = "15GSBB"
        GSBC.Text = "15GSBC"
        GSBC_1.Text = "15GSBC"
        GSBC_2.Text = "15GSBC"
        GSBC_3.Text = "15GSBC"
        GSBC_4.Text = "15GSBC"
        GSBC_5.Text = "15GSBC"
        SPA.Text = "15MVSA"
        spa_3.Text = "15MVSA"
        SPA_2.Text = "15MVSA"
        SPB.Text = "15MVSB"
        SPB_2.Text = "15MVSB"
        SPB_3.Text = "15MVSB"
        spc_1.Text = "15MVSC"
        SPC_3.Text = "15MVSC"
        SPC_4.Text = "15MVSC"
    End If
'If LAS 9 Power System 2 is selected
    If cboPS.value = "Power System 2  " Then
        BS1.Text = "16S-BS1"
        BS1_1.Text = "16S-BS1"
        GSBA.Text = "16GSBA"
        GSBA_1.Text = "16GSBA"
        GSBA_2.Text = "16GSBA"
        GSBA_3.Text = "16GSBA"
        GSBA_4.Text = "16GSBA"
        GSBA_5.Text = "16GSBA"
        GSBB.Text = "16GSBB"
        GSBB_1.Text = "16GSBB"
        GSBB_2.Text = "16GSBB"
        GSBB_3.Text = "16GSBB"
        GSBB_4.Text = "16GSBB"
        GSBB_5.Text = "16GSBB"
        GSBB_6.Text = "16GSBB"
        GSBC.Text = "16GSBC"
        GSBC_1.Text = "16GSBC"
        GSBC_2.Text = "16GSBC"
        GSBC_3.Text = "16GSBC"
        GSBC_4.Text = "16GSBC"
        GSBC_5.Text = "16GSBC"
        SPA.Text = "16MVSA"
        spa_3.Text = "16MVSA"
        SPA_2.Text = "16MVSA"
        SPB.Text = "16MVSB"
        SPB_2.Text = "16MVSB"
        SPB_3.Text = "16MVSB"
        spc_1.Text = "16MVSC"
        SPC_3.Text = "16MVSC"
        SPC_4.Text = "16MVSC"
    End If
'If LAS 9 Power System 3 is selected
    If cboPS.value = "Power System 3  " Then
        BS1.Text = "17S-BS1"
        BS1_1.Text = "17S-BS1"
        GSBA.Text = "17GSBA"
        GSBA_1.Text = "17GSBA"
        GSBA_2.Text = "17GSBA"
        GSBA_3.Text = "17GSBA"
        GSBA_4.Text = "17GSBA"
        GSBA_5.Text = "17GSBA"
        GSBB.Text = "17GSBB"
        GSBB_1.Text = "17GSBB"
        GSBB_2.Text = "17GSBB"
        GSBB_3.Text = "17GSBB"
        GSBB_4.Text = "17GSBB"
        GSBB_5.Text = "17GSBB"
        GSBB_6.Text = "17GSBB"
        GSBC.Text = "17GSBC"
        GSBC_1.Text = "17GSBC"
        GSBC_2.Text = "17GSBC"
        GSBC_3.Text = "17GSBC"
        GSBC_4.Text = "17GSBC"
        GSBC_5.Text = "17GSBC"
        SPA.Text = "17MVSA"
        spa_3.Text = "17MVSA"
        SPA_2.Text = "17MVSA"
        SPB.Text = "17MVSB"
        SPB_2.Text = "17MVSB"
        SPB_3.Text = "17MVSB"
        spc_1.Text = "17MVSC"
        SPC_3.Text = "17MVSC"
        SPC_4.Text = "17MVSC"
    End If
'If LAS 9 Power System 4 is selected
    If cboPS.value = "Power System 4  " Then
        BS1.Text = "18S-BS1"
        BS1_1.Text = "18S-BS1"
        GSBA.Text = "18GSBA"
        GSBA_1.Text = "18GSBA"
        GSBA_2.Text = "18GSBA"
        GSBA_3.Text = "18GSBA"
        GSBA_4.Text = "18GSBA"
        GSBA_5.Text = "18GSBA"
        GSBB.Text = "18GSBB"
        GSBB_1.Text = "18GSBB"
        GSBB_2.Text = "18GSBB"
        GSBB_3.Text = "18GSBB"
        GSBB_4.Text = "18GSBB"
        GSBB_5.Text = "18GSBB"
        GSBB_6.Text = "18GSBB"
        GSBC.Text = "18GSBC"
        GSBC_1.Text = "18GSBC"
        GSBC_2.Text = "18GSBC"
        GSBC_3.Text = "18GSBC"
        GSBC_4.Text = "18GSBC"
        GSBC_5.Text = "18GSBC"
        SPA.Text = "18MVSA"
        spa_3.Text = "18MVSA"
        SPA_2.Text = "18MVSA"
        SPB.Text = "18MVSB"
        SPB_2.Text = "18MVSB"
        SPB_3.Text = "18MVSB"
        spc_1.Text = "18MVSC"
        SPC_3.Text = "18MVSC"
        SPC_4.Text = "18MVSC"
    End If
'If LAS 10 Power System 1 is selected
    If cboPS.value = " Power System 1" Then
        BS1.Text = "1S-BS1"
        BS1_1.Text = "1S-BS1"
        GSBA.Text = "1GSBA"
        GSBA_1.Text = "1GSBA"
        GSBA_2.Text = "1GSBA"
        GSBA_3.Text = "1GSBA"
        GSBA_4.Text = "1GSBA"
        GSBA_5.Text = "1GSBA"
        GSBB.Text = "1GSBB"
        GSBB_1.Text = "1GSBB"
        GSBB_2.Text = "1GSBB"
        GSBB_3.Text = "1GSBB"
        GSBB_4.Text = "1GSBB"
        GSBB_5.Text = "1GSBB"
        GSBB_6.Text = "1GSBB"
        GSBC.Text = "1GSBC"
        GSBC_1.Text = "1GSBC"
        GSBC_2.Text = "1GSBC"
        GSBC_3.Text = "1GSBC"
        GSBC_4.Text = "1GSBC"
        GSBC_5.Text = "1GSBC"
        SPA.Text = "1MVSA"
        spa_3.Text = "1MVSA"
        SPA_2.Text = "1MVSA"
        SPB.Text = "1MVSB"
        SPB_2.Text = "1MVSB"
        SPB_3.Text = "1MVSB"
        spc_1.Text = "1MVSC"
        SPC_3.Text = "1MVSC"
        SPC_4.Text = "1MVSC"
    End If
'If LAS 10 Power System 2 is selected
    If cboPS.value = " Power System 2" Then
        BS1.Text = "2S-BS1"
        BS1_1.Text = "2S-BS1"
        GSBA.Text = "2GSBA"
        GSBA_1.Text = "2GSBA"
        GSBA_2.Text = "2GSBA"
        GSBA_3.Text = "2GSBA"
        GSBA_4.Text = "2GSBA"
        GSBA_5.Text = "2GSBA"
        GSBB.Text = "2GSBB"
        GSBB_1.Text = "2GSBB"
        GSBB_2.Text = "2GSBB"
        GSBB_3.Text = "2GSBB"
        GSBB_4.Text = "2GSBB"
        GSBB_5.Text = "2GSBB"
        GSBB_6.Text = "2GSBB"
        GSBC.Text = "2GSBC"
        GSBC_1.Text = "2GSBC"
        GSBC_2.Text = "2GSBC"
        GSBC_3.Text = "2GSBC"
        GSBC_4.Text = "2GSBC"
        GSBC_5.Text = "2GSBC"
        SPA.Text = "2MVSA"
        spa_3.Text = "2MVSA"
        SPA_2.Text = "2MVSA"
        SPB.Text = "2MVSB"
        SPB_2.Text = "2MVSB"
        SPB_3.Text = "2MVSB"
        spc_1.Text = "2MVSC"
        SPC_3.Text = "2MVSC"
        SPC_4.Text = "2MVSC"
    End If
'If LAS 10 Power System 3 is selected
    If cboPS.value = " Power System 3" Then
        BS1.Text = "3S-BS1"
        BS1_1.Text = "3S-BS1"
        GSBA.Text = "3GSBA"
        GSBA_1.Text = "3GSBA"
        GSBA_2.Text = "3GSBA"
        GSBA_3.Text = "3GSBA"
        GSBA_4.Text = "3GSBA"
        GSBA_5.Text = "3GSBA"
        GSBB.Text = "3GSBB"
        GSBB_1.Text = "3GSBB"
        GSBB_2.Text = "3GSBB"
        GSBB_3.Text = "3GSBB"
        GSBB_4.Text = "3GSBB"
        GSBB_5.Text = "3GSBB"
        GSBB_6.Text = "3GSBB"
        GSBC.Text = "3GSBC"
        GSBC_1.Text = "3GSBC"
        GSBC_2.Text = "3GSBC"
        GSBC_3.Text = "3GSBC"
        GSBC_4.Text = "3GSBC"
        GSBC_5.Text = "3GSBC"
        SPA.Text = "3MVSA"
        spa_3.Text = "3MVSA"
        SPA_2.Text = "3MVSA"
        SPB.Text = "3MVSB"
        SPB_2.Text = "3MVSB"
        SPB_3.Text = "3MVSB"
        spc_1.Text = "3MVSC"
        SPC_3.Text = "3MVSC"
        SPC_4.Text = "3MVSC"
        
    End If
'If LAS 11 Power System 1 is selected
    If cboPS.value = "Power System 1" Then
        BS1.Text = "1S-BS1"
        BS1_1.Text = "1S-BS1"
        GSBA.Text = "1GSBA"
        GSBA_1.Text = "1GSBA"
        GSBA_2.Text = "1GSBA"
        GSBA_3.Text = "1GSBA"
        GSBA_4.Text = "1GSBA"
        GSBA_5.Text = "1GSBA"
        GSBB.Text = "1GSBB"
        GSBB_1.Text = "1GSBB"
        GSBB_2.Text = "1GSBB"
        GSBB_3.Text = "1GSBB"
        GSBB_4.Text = "1GSBB"
        GSBB_5.Text = "1GSBB"
        GSBB_6.Text = "1GSBB"
        GSBC.Text = "1GSBC"
        GSBC_1.Text = "1GSBC"
        GSBC_2.Text = "1GSBC"
        GSBC_3.Text = "1GSBC"
        GSBC_4.Text = "1GSBC"
        GSBC_5.Text = "1GSBC"
        SPA.Text = "1MVSA"
        spa_3.Text = "1MVSA"
        SPA_2.Text = "1MVSA"
        SPB.Text = "1MVSB"
        SPB_2.Text = "1MVSB"
        SPB_3.Text = "1MVSB"
        spc_1.Text = "1MVSC"
        SPC_3.Text = "1MVSC"
        SPC_4.Text = "1MVSC"
    End If
'If LAS 11 Power System 2 is selected
    If cboPS.value = "Power System 2" Then
        BS1.Text = "2S-BS1"
        BS1_1.Text = "2S-BS1"
        GSBA.Text = "2GSBA"
        GSBA_1.Text = "2GSBA"
        GSBA_2.Text = "2GSBA"
        GSBA_3.Text = "2GSBA"
        GSBA_4.Text = "2GSBA"
        GSBA_5.Text = "2GSBA"
        GSBB.Text = "2GSBB"
        GSBB_1.Text = "2GSBB"
        GSBB_2.Text = "2GSBB"
        GSBB_3.Text = "2GSBB"
        GSBB_4.Text = "2GSBB"
        GSBB_5.Text = "2GSBB"
        GSBB_6.Text = "2GSBB"
        GSBC.Text = "2GSBC"
        GSBC_1.Text = "2GSBC"
        GSBC_2.Text = "2GSBC"
        GSBC_3.Text = "2GSBC"
        GSBC_4.Text = "2GSBC"
        GSBC_5.Text = "2GSBC"
        SPA.Text = "2MVSA"
        spa_3.Text = "2MVSA"
        SPA_2.Text = "2MVSA"
        SPB.Text = "2MVSB"
        SPB_2.Text = "2MVSB"
        SPB_3.Text = "2MVSB"
        spc_1.Text = "2MVSC"
        SPC_3.Text = "2MVSC"
        SPC_4.Text = "2MVSC"
    End If
'If LAS 11 Power System 3 is selected
    If cboPS.value = "Power System 3" Then
        BS1.Text = "3S-BS1"
        BS1_1.Text = "3S-BS1"
        GSBA.Text = "3GSBA"
        GSBA_1.Text = "3GSBA"
        GSBA_2.Text = "3GSBA"
        GSBA_3.Text = "3GSBA"
        GSBA_4.Text = "3GSBA"
        GSBA_5.Text = "3GSBA"
        GSBB.Text = "3GSBB"
        GSBB_1.Text = "3GSBB"
        GSBB_2.Text = "3GSBB"
        GSBB_3.Text = "3GSBB"
        GSBB_4.Text = "3GSBB"
        GSBB_5.Text = "3GSBB"
        GSBB_6.Text = "3GSBB"
        GSBC.Text = "3GSBC"
        GSBC_1.Text = "3GSBC"
        GSBC_2.Text = "3GSBC"
        GSBC_3.Text = "3GSBC"
        GSBC_4.Text = "3GSBC"
        GSBC_5.Text = "3GSBC"
        SPA.Text = "3MVSA"
        spa_3.Text = "3MVSA"
        SPA_2.Text = "3MVSA"
        SPB.Text = "3MVSB"
        SPB_2.Text = "3MVSB"
        SPB_3.Text = "3MVSB"
        spc_1.Text = "3MVSC"
        SPC_3.Text = "3MVSC"
        SPC_4.Text = "3MVSC"
    End If
        
End Sub
Private Sub generatorselection_1()

'Set generators when the load bank is not involved.
    
    Dim Gen_1 As Range
    Dim Gen_1_2 As Range
    Dim Gen_1_3 As Range
    Dim Gen_1_6 As Range
    Dim Gen_1_7 As Range
    Dim Gen_2 As Range
    Dim Gen_2_2 As Range
    Dim Gen_2_3 As Range
    Dim Gen_2_6 As Range
    Dim Gen_2_7 As Range
    Dim Gen_3 As Range
    Dim Gen_3_2 As Range
    Dim Gen_3_3 As Range
    Dim Gen_3_6 As Range
    Dim Gen_3_7 As Range
    Dim Gen_4 As Range
    Dim Gen_4_2 As Range
    Dim Gen_4_3 As Range
    Dim Gen_4_6 As Range
    Dim Gen_4_7 As Range
    Dim Gen_5 As Range
    Dim Gen_5_2 As Range
    Dim Gen_5_3 As Range
    Dim Gen_5_6 As Range
    Dim Gen_5_7 As Range
    Dim Gen_6 As Range
    Dim Gen_6_2 As Range
    Dim Gen_6_3 As Range
    Dim Gen_6_6 As Range
    Dim Gen_6_7 As Range
    
If cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11" Then
    Set Gen_1 = ActiveDocument.Bookmarks("tGen_1").Range
    Set Gen_2 = ActiveDocument.Bookmarks("tGen_2").Range
    Set Gen_3 = ActiveDocument.Bookmarks("tGen_3").Range
    Set Gen_4 = ActiveDocument.Bookmarks("tGen_4").Range
    Set Gen_5 = ActiveDocument.Bookmarks("tGen_5").Range
    Set Gen_6 = ActiveDocument.Bookmarks("tGen_6").Range
    Set Gen_1_2 = ActiveDocument.Bookmarks("tGen_1_2").Range
    Set Gen_1_3 = ActiveDocument.Bookmarks("tGen_1_3").Range
    Set Gen_1_6 = ActiveDocument.Bookmarks("tGen_1_6").Range
    Set Gen_1_7 = ActiveDocument.Bookmarks("tGen_1_7").Range
    Set Gen_2_2 = ActiveDocument.Bookmarks("tGen_2_2").Range
    Set Gen_2_3 = ActiveDocument.Bookmarks("tGen_2_3").Range
    Set Gen_2_6 = ActiveDocument.Bookmarks("tGen_2_6").Range
    Set Gen_2_7 = ActiveDocument.Bookmarks("tGen_2_7").Range
    Set Gen_3_2 = ActiveDocument.Bookmarks("tGen_3_2").Range
    Set Gen_3_3 = ActiveDocument.Bookmarks("tGen_3_3").Range
    Set Gen_3_6 = ActiveDocument.Bookmarks("tGen_3_6").Range
    Set Gen_3_7 = ActiveDocument.Bookmarks("tGen_3_7").Range
    Set Gen_4_2 = ActiveDocument.Bookmarks("tGen_4_2").Range
    Set Gen_4_3 = ActiveDocument.Bookmarks("tGen_4_3").Range
    Set Gen_4_6 = ActiveDocument.Bookmarks("tGen_4_6").Range
    Set Gen_4_7 = ActiveDocument.Bookmarks("tGen_4_7").Range
    Set Gen_5_2 = ActiveDocument.Bookmarks("tGen_5_2").Range
    Set Gen_5_3 = ActiveDocument.Bookmarks("tGen_5_3").Range
    Set Gen_5_6 = ActiveDocument.Bookmarks("tGen_5_6").Range
    Set Gen_5_7 = ActiveDocument.Bookmarks("tGen_5_7").Range
    Set Gen_6_2 = ActiveDocument.Bookmarks("tGen_6_2").Range
    Set Gen_6_3 = ActiveDocument.Bookmarks("tGen_6_3").Range
    Set Gen_6_6 = ActiveDocument.Bookmarks("tGen_6_6").Range
    Set Gen_6_7 = ActiveDocument.Bookmarks("tGen_6_7").Range
End If

    If cboPS.value = "Power System 11" Then
            Gen_1.Text = "11-1"
            Gen_1_2.Text = "11-1"
            Gen_1_3.Text = "11-1"
            Gen_1_6.Text = "11-1"
            Gen_1_7.Text = "11-1"
            Gen_2.Text = "11-2"
            Gen_2_2.Text = "11-2"
            Gen_2_3.Text = "11-2"
            Gen_2_6.Text = "11-2"
            Gen_2_7.Text = "11-2"
            Gen_3.Text = "11-3"
            Gen_3_2.Text = "11-3"
            Gen_3_3.Text = "11-3"
            Gen_3_6.Text = "11-3"
            Gen_3_7.Text = "11-3"
            Gen_4.Text = "11-4"
            Gen_4_2.Text = "11-4"
            Gen_4_3.Text = "11-4"
            Gen_4_6.Text = "11-4"
            Gen_4_7.Text = "11-4"
            Gen_5.Text = "11-5"
            Gen_5_2.Text = "11-5"
            Gen_5_3.Text = "11-5"
            Gen_5_6.Text = "11-5"
            Gen_5_7.Text = "11-5"
            Gen_6.Text = "11-6"
            Gen_6_2.Text = "11-6"
            Gen_6_3.Text = "11-6"
            Gen_6_6.Text = "11-6"
            Gen_6_7.Text = "11-6"
        ElseIf cboPS.value = "Power System 12" Then
            Gen_1.Text = "12-1"
            Gen_1_2.Text = "12-1"
            Gen_1_3.Text = "12-1"
            Gen_1_6.Text = "12-1"
            Gen_1_7.Text = "12-1"
            Gen_2.Text = "12-2"
            Gen_2_2.Text = "12-2"
            Gen_2_3.Text = "12-2"
            Gen_2_6.Text = "12-2"
            Gen_2_7.Text = "12-2"
            Gen_3.Text = "12-3"
            Gen_3_2.Text = "12-3"
            Gen_3_3.Text = "12-3"
            Gen_3_6.Text = "12-3"
            Gen_3_7.Text = "12-3"
            Gen_4.Text = "12-4"
            Gen_4_2.Text = "12-4"
            Gen_4_3.Text = "12-4"
            Gen_4_6.Text = "12-4"
            Gen_4_7.Text = "12-4"
            Gen_5.Text = "12-5"
            Gen_5_2.Text = "12-5"
            Gen_5_3.Text = "12-5"
            Gen_5_6.Text = "12-5"
            Gen_5_7.Text = "12-5"
            Gen_6.Text = "12-6"
            Gen_6_2.Text = "12-6"
            Gen_6_3.Text = "12-6"
            Gen_6_6.Text = "12-6"
            Gen_6_7.Text = "12-6"
        ElseIf cboPS.value = "Power System 13" Then
            Gen_1.Text = "13-1"
            Gen_1_2.Text = "13-1"
            Gen_1_3.Text = "13-1"
            Gen_1_6.Text = "13-1"
            Gen_1_7.Text = "13-1"
            Gen_2.Text = "13-2"
            Gen_2_2.Text = "13-2"
            Gen_2_3.Text = "13-2"
            Gen_2_6.Text = "13-2"
            Gen_2_7.Text = "13-2"
            Gen_3.Text = "13-3"
            Gen_3_2.Text = "13-3"
            Gen_3_3.Text = "13-3"
            Gen_3_6.Text = "13-3"
            Gen_3_7.Text = "13-3"
            Gen_4.Text = "13-4"
            Gen_4_2.Text = "13-4"
            Gen_4_3.Text = "13-4"
            Gen_4_6.Text = "13-4"
            Gen_4_7.Text = "13-4"
            Gen_5.Text = "13-5"
            Gen_5_2.Text = "13-5"
            Gen_5_3.Text = "13-5"
            Gen_5_6.Text = "13-5"
            Gen_5_7.Text = "13-5"
            Gen_6.Text = "13-6"
            Gen_6_2.Text = "13-6"
            Gen_6_3.Text = "13-6"
            Gen_6_6.Text = "13-6"
            Gen_6_7.Text = "13-6"
        ElseIf cboPS.value = "Power System 14" Then
            Gen_1.Text = "14-1"
            Gen_1_2.Text = "14-1"
            Gen_1_3.Text = "14-1"
            Gen_1_6.Text = "14-1"
            Gen_1_7.Text = "14-1"
            Gen_2.Text = "14-2"
            Gen_2_2.Text = "14-2"
            Gen_2_3.Text = "14-2"
            Gen_2_6.Text = "14-2"
            Gen_2_7.Text = "14-2"
            Gen_3.Text = "14-3"
            Gen_3_2.Text = "14-3"
            Gen_3_3.Text = "14-3"
            Gen_3_6.Text = "14-3"
            Gen_3_7.Text = "14-3"
            Gen_4.Text = "14-4"
            Gen_4_2.Text = "14-4"
            Gen_4_3.Text = "14-4"
            Gen_4_6.Text = "14-4"
            Gen_4_7.Text = "14-4"
            Gen_5.Text = "14-5"
            Gen_5_2.Text = "14-5"
            Gen_5_3.Text = "14-5"
            Gen_5_6.Text = "14-5"
            Gen_5_7.Text = "14-5"
            Gen_6.Text = "14-6"
            Gen_6_2.Text = "14-6"
            Gen_6_3.Text = "14-6"
            Gen_6_6.Text = "14-6"
            Gen_6_7.Text = "14-6"
        ElseIf cboPS.value = "Power System 1  " Then
            Gen_1.Text = "1-1"
            Gen_1_2.Text = "1-1"
            Gen_1_3.Text = "1-1"
            Gen_1_6.Text = "1-1"
            Gen_1_7.Text = "1-1"
            Gen_2.Text = "1-2"
            Gen_2_2.Text = "1-2"
            Gen_2_3.Text = "1-2"
            Gen_2_6.Text = "1-2"
            Gen_2_7.Text = "1-2"
            Gen_3.Text = "1-3"
            Gen_3_2.Text = "1-3"
            Gen_3_3.Text = "1-3"
            Gen_3_6.Text = "1-3"
            Gen_3_7.Text = "1-3"
            Gen_4.Text = "1-4"
            Gen_4_2.Text = "1-4"
            Gen_4_3.Text = "1-4"
            Gen_4_6.Text = "1-4"
            Gen_4_7.Text = "1-4"
            Gen_5.Text = "1-5"
            Gen_5_2.Text = "1-5"
            Gen_5_3.Text = "1-5"
            Gen_5_6.Text = "1-5"
            Gen_5_7.Text = "1-5"
            Gen_6.Text = "1-6"
            Gen_6_2.Text = "1-6"
            Gen_6_3.Text = "1-6"
            Gen_6_6.Text = "1-6"
            Gen_6_7.Text = "1-6"
        ElseIf cboPS.value = "Power System 2  " Then
            Gen_1.Text = "2-1"
            Gen_1_2.Text = "2-1"
            Gen_1_3.Text = "2-1"
            Gen_1_6.Text = "2-1"
            Gen_1_7.Text = "2-1"
            Gen_2.Text = "2-2"
            Gen_2_2.Text = "2-2"
            Gen_2_3.Text = "2-2"
            Gen_2_6.Text = "2-2"
            Gen_2_7.Text = "2-2"
            Gen_3.Text = "2-3"
            Gen_3_2.Text = "2-3"
            Gen_3_3.Text = "2-3"
            Gen_3_6.Text = "2-3"
            Gen_3_7.Text = "2-3"
            Gen_4.Text = "2-4"
            Gen_4_2.Text = "2-4"
            Gen_4_3.Text = "2-4"
            Gen_4_6.Text = "2-4"
            Gen_4_7.Text = "2-4"
            Gen_5.Text = "2-5"
            Gen_5_2.Text = "2-5"
            Gen_5_3.Text = "2-5"
            Gen_5_6.Text = "2-5"
            Gen_5_7.Text = "2-5"
            Gen_6.Text = "2-6"
            Gen_6_2.Text = "2-6"
            Gen_6_3.Text = "2-6"
            Gen_6_6.Text = "2-6"
            Gen_6_7.Text = "2-6"
        ElseIf cboPS.value = "Power System 3  " Then
            Gen_1.Text = "3-1"
            Gen_1_2.Text = "3-1"
            Gen_1_3.Text = "3-1"
            Gen_1_6.Text = "3-1"
            Gen_1_7.Text = "3-1"
            Gen_2.Text = "3-2"
            Gen_2_2.Text = "3-2"
            Gen_2_3.Text = "3-2"
            Gen_2_6.Text = "3-2"
            Gen_2_7.Text = "3-2"
            Gen_3.Text = "3-3"
            Gen_3_2.Text = "3-3"
            Gen_3_3.Text = "3-3"
            Gen_3_6.Text = "3-3"
            Gen_3_7.Text = "3-3"
            Gen_4.Text = "3-4"
            Gen_4_2.Text = "3-4"
            Gen_4_3.Text = "3-4"
            Gen_4_6.Text = "3-4"
            Gen_4_7.Text = "3-4"
            Gen_5.Text = "3-5"
            Gen_5_2.Text = "3-5"
            Gen_5_3.Text = "3-5"
            Gen_5_6.Text = "3-5"
            Gen_5_7.Text = "3-5"
            Gen_6.Text = "3-6"
            Gen_6_2.Text = "3-6"
            Gen_6_3.Text = "3-6"
            Gen_6_6.Text = "3-6"
            Gen_6_7.Text = "3-6"
        ElseIf cboPS.value = "Power System 4  " Then
            Gen_1.Text = "4-1"
            Gen_1_2.Text = "4-1"
            Gen_1_3.Text = "4-1"
            Gen_1_6.Text = "4-1"
            Gen_1_7.Text = "4-1"
            Gen_2.Text = "4-2"
            Gen_2_2.Text = "4-2"
            Gen_2_3.Text = "4-2"
            Gen_2_6.Text = "4-2"
            Gen_2_7.Text = "4-2"
            Gen_3.Text = "4-3"
            Gen_3_2.Text = "4-3"
            Gen_3_3.Text = "4-3"
            Gen_3_6.Text = "4-3"
            Gen_3_7.Text = "4-3"
            Gen_4.Text = "4-4"
            Gen_4_2.Text = "4-4"
            Gen_4_3.Text = "4-4"
            Gen_4_6.Text = "4-4"
            Gen_4_7.Text = "4-4"
            Gen_5.Text = "4-5"
            Gen_5_2.Text = "4-5"
            Gen_5_3.Text = "4-5"
            Gen_5_6.Text = "4-5"
            Gen_5_7.Text = "4-5"
            Gen_6.Text = "4-6"
            Gen_6_2.Text = "4-6"
            Gen_6_3.Text = "4-6"
            Gen_6_6.Text = "4-6"
            Gen_6_7.Text = "4-6"
        ElseIf cboPS.value = " Power System 1" Then
            Gen_1.Text = "1-1"
            Gen_1_2.Text = "1-1"
            Gen_1_3.Text = "1-1"
            Gen_1_6.Text = "1-1"
            Gen_1_7.Text = "1-1"
            Gen_2.Text = "1-2"
            Gen_2_2.Text = "1-2"
            Gen_2_3.Text = "1-2"
            Gen_2_6.Text = "1-2"
            Gen_2_7.Text = "1-2"
            Gen_3.Text = "1-3"
            Gen_3_2.Text = "1-3"
            Gen_3_3.Text = "1-3"
            Gen_3_6.Text = "1-3"
            Gen_3_7.Text = "1-3"
            Gen_4.Text = "1-4"
            Gen_4_2.Text = "1-4"
            Gen_4_3.Text = "1-4"
            Gen_4_6.Text = "1-4"
            Gen_4_7.Text = "1-4"
            Gen_5.Text = "1-5"
            Gen_5_2.Text = "1-5"
            Gen_5_3.Text = "1-5"
            Gen_5_6.Text = "1-5"
            Gen_5_7.Text = "1-5"
            Gen_6.Text = "1-6"
            Gen_6_2.Text = "1-6"
            Gen_6_3.Text = "1-6"
            Gen_6_6.Text = "1-6"
            Gen_6_7.Text = "1-6"
        ElseIf cboPS.value = " Power System 2" Then
            Gen_1.Text = "2-1"
            Gen_1_2.Text = "2-1"
            Gen_1_3.Text = "2-1"
            Gen_1_6.Text = "2-1"
            Gen_1_7.Text = "2-1"
            Gen_2.Text = "2-2"
            Gen_2_2.Text = "2-2"
            Gen_2_3.Text = "2-2"
            Gen_2_6.Text = "2-2"
            Gen_2_7.Text = "2-2"
            Gen_3.Text = "2-3"
            Gen_3_2.Text = "2-3"
            Gen_3_3.Text = "2-3"
            Gen_3_6.Text = "2-3"
            Gen_3_7.Text = "2-3"
            Gen_4.Text = "2-4"
            Gen_4_2.Text = "2-4"
            Gen_4_3.Text = "2-4"
            Gen_4_6.Text = "2-4"
            Gen_4_7.Text = "2-4"
            Gen_5.Text = "2-5"
            Gen_5_2.Text = "2-5"
            Gen_5_3.Text = "2-5"
            Gen_5_6.Text = "2-5"
            Gen_5_7.Text = "2-5"
            Gen_6.Text = "2-6"
            Gen_6_2.Text = "2-6"
            Gen_6_3.Text = "2-6"
            Gen_6_6.Text = "2-6"
            Gen_6_7.Text = "2-6"
        ElseIf cboPS.value = " Power System 3" Then
            Gen_1.Text = "3-1"
            Gen_1_2.Text = "3-1"
            Gen_1_3.Text = "3-1"
            Gen_1_6.Text = "3-1"
            Gen_1_7.Text = "3-1"
            Gen_2.Text = "3-2"
            Gen_2_2.Text = "3-2"
            Gen_2_3.Text = "3-2"
            Gen_2_6.Text = "3-2"
            Gen_2_7.Text = "3-2"
            Gen_3.Text = "3-3"
            Gen_3_2.Text = "3-3"
            Gen_3_3.Text = "3-3"
            Gen_3_6.Text = "3-3"
            Gen_3_7.Text = "3-3"
            Gen_4.Text = "11-4"
            Gen_4_2.Text = "3-4"
            Gen_4_3.Text = "3-4"
            Gen_4_6.Text = "3-4"
            Gen_4_7.Text = "3-4"
            Gen_5.Text = "3-5"
            Gen_5_2.Text = "3-5"
            Gen_5_3.Text = "3-5"
            Gen_5_6.Text = "3-5"
            Gen_5_7.Text = "3-5"
            Gen_6.Text = "3-6"
            Gen_6_2.Text = "3-6"
            Gen_6_3.Text = "3-6"
            Gen_6_6.Text = "3-6"
            Gen_6_7.Text = "3-6"
        ElseIf cboPS.value = "Power System 1" Then
            Gen_1.Text = "1-1"
            Gen_1_2.Text = "1-1"
            Gen_1_3.Text = "1-1"
            Gen_1_6.Text = "1-1"
            Gen_1_7.Text = "1-1"
            Gen_2.Text = "1-2"
            Gen_2_2.Text = "1-2"
            Gen_2_3.Text = "1-2"
            Gen_2_6.Text = "1-2"
            Gen_2_7.Text = "1-2"
            Gen_3.Text = "1-3"
            Gen_3_2.Text = "1-3"
            Gen_3_3.Text = "1-3"
            Gen_3_6.Text = "1-3"
            Gen_3_7.Text = "1-3"
            Gen_4.Text = "1-4"
            Gen_4_2.Text = "1-4"
            Gen_4_3.Text = "1-4"
            Gen_4_6.Text = "1-4"
            Gen_4_7.Text = "1-4"
            Gen_5.Text = "1-5"
            Gen_5_2.Text = "1-5"
            Gen_5_3.Text = "1-5"
            Gen_5_6.Text = "1-5"
            Gen_5_7.Text = "1-5"
            Gen_6.Text = "1-6"
            Gen_6_2.Text = "1-6"
            Gen_6_3.Text = "1-6"
            Gen_6_6.Text = "1-6"
            Gen_6_7.Text = "1-6"
        ElseIf cboPS.value = "Power System 2" Then
            Gen_1.Text = "2-1"
            Gen_1_2.Text = "2-1"
            Gen_1_3.Text = "2-1"
            Gen_1_6.Text = "2-1"
            Gen_1_7.Text = "2-1"
            Gen_2.Text = "2-2"
            Gen_2_2.Text = "2-2"
            Gen_2_3.Text = "2-2"
            Gen_2_6.Text = "2-2"
            Gen_2_7.Text = "2-2"
            Gen_3.Text = "2-3"
            Gen_3_2.Text = "2-3"
            Gen_3_3.Text = "2-3"
            Gen_3_6.Text = "2-3"
            Gen_3_7.Text = "2-3"
            Gen_4.Text = "2-4"
            Gen_4_2.Text = "2-4"
            Gen_4_3.Text = "2-4"
            Gen_4_6.Text = "2-4"
            Gen_4_7.Text = "2-4"
            Gen_5.Text = "2-5"
            Gen_5_2.Text = "2-5"
            Gen_5_3.Text = "2-5"
            Gen_5_6.Text = "2-5"
            Gen_5_7.Text = "2-5"
            Gen_6.Text = "2-6"
            Gen_6_2.Text = "2-6"
            Gen_6_3.Text = "2-6"
            Gen_6_6.Text = "2-6"
            Gen_6_7.Text = "2-6"
        ElseIf cboPS.value = "Power System 3" Then
            Gen_1.Text = "3-1"
            Gen_1_2.Text = "3-1"
            Gen_1_3.Text = "3-1"
            Gen_1_6.Text = "3-1"
            Gen_1_7.Text = "3-1"
            Gen_2.Text = "3-2"
            Gen_2_2.Text = "3-2"
            Gen_2_3.Text = "3-2"
            Gen_2_6.Text = "3-2"
            Gen_2_7.Text = "3-2"
            Gen_3.Text = "3-3"
            Gen_3_2.Text = "3-3"
            Gen_3_3.Text = "3-3"
            Gen_3_6.Text = "3-3"
            Gen_3_7.Text = "3-3"
            Gen_4.Text = "3-4"
            Gen_4_2.Text = "3-4"
            Gen_4_3.Text = "3-4"
            Gen_4_6.Text = "3-4"
            Gen_4_7.Text = "3-4"
            Gen_5.Text = "3-5"
            Gen_5_2.Text = "3-5"
            Gen_5_3.Text = "3-5"
            Gen_5_6.Text = "3-5"
            Gen_5_7.Text = "3-5"
            Gen_6.Text = "3-6"
            Gen_6_2.Text = "3-6"
            Gen_6_3.Text = "3-6"
            Gen_6_6.Text = "3-6"
            Gen_6_7.Text = "3-6"
    End If

    If cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = "1 UPS Corrective Maintenance" _
    Or cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Then
        Set GSBA_1 = Nothing
        Set GSBC_1 = Nothing
        Set GSBB_2 = Nothing
        Set GSBB_1 = Nothing
        Set BS1 = ActiveDocument.Bookmarks("tBS1").Range
        Set BS1_1 = ActiveDocument.Bookmarks("tBS1_1").Range
        Set GSBA = ActiveDocument.Bookmarks("tGSBA").Range
        Set GSBA_2 = ActiveDocument.Bookmarks("tGSBA_2").Range
        Set GSBA_3 = ActiveDocument.Bookmarks("tGSBA_3").Range
        Set GSBA_4 = ActiveDocument.Bookmarks("tGSBA_4").Range
        Set GSBA_5 = ActiveDocument.Bookmarks("tGSBA_5").Range
        Set GSBB = ActiveDocument.Bookmarks("tGSBB").Range
        Set GSBB_3 = ActiveDocument.Bookmarks("tGSBB_3").Range
        Set GSBB_4 = ActiveDocument.Bookmarks("tGSBB_4").Range
        Set GSBB_5 = ActiveDocument.Bookmarks("tGSBB_5").Range
        Set GSBB_6 = ActiveDocument.Bookmarks("tGSBB_6").Range
        Set GSBC = ActiveDocument.Bookmarks("tGSBC").Range
        Set GSBC_2 = ActiveDocument.Bookmarks("tGSBC_2").Range
        Set GSBC_3 = ActiveDocument.Bookmarks("tGSBC_3").Range
        Set GSBC_4 = ActiveDocument.Bookmarks("tGSBC_4").Range
        Set GSBC_5 = ActiveDocument.Bookmarks("tGSBC_5").Range
        Set SPA = ActiveDocument.Bookmarks("tSPA").Range
        Set spa_3 = ActiveDocument.Bookmarks("tspa_1").Range
        Set SPA_2 = ActiveDocument.Bookmarks("tSPA_2").Range
        Set SPB = ActiveDocument.Bookmarks("tSPB").Range
        Set SPB_2 = ActiveDocument.Bookmarks("tSPB_2").Range
        Set SPB_3 = ActiveDocument.Bookmarks("tSPB_3").Range
        Set spc_1 = ActiveDocument.Bookmarks("tspc_1").Range
        Set SPC_3 = ActiveDocument.Bookmarks("tSPC_3").Range
        Set SPC_4 = ActiveDocument.Bookmarks("tSPC_4").Range
   
    End If

'If Power System 11 is selected
    If cboPS.value = "Power System 11" Then
        BS1.Text = "11S-BS1"
        BS1_1.Text = "11S-BS1"
        GSBA.Text = "11GSBA"
        GSBA_2.Text = "11GSBA"
        GSBA_3.Text = "11GSBA"
        GSBA_4.Text = "11GSBA"
        GSBA_5.Text = "11GSBA"
        GSBB.Text = "11GSBB"
        GSBB_3.Text = "11GSBB"
        GSBB_4.Text = "11GSBB"
        GSBB_5.Text = "11GSBB"
        GSBB_6.Text = "11GSBB"
        GSBC.Text = "11GSBC"
        GSBC_2.Text = "11GSBC"
        GSBC_3.Text = "11GSBC"
        GSBC_4.Text = "11GSBC"
        GSBC_5.Text = "11GSBC"
        SPA.Text = "11MVSA"
        spa_3.Text = "11MVSA"
        SPA_2.Text = "11MVSA"
        SPB.Text = "11MVSB"
        SPB_2.Text = "11MVSB"
        SPB_3.Text = "11MVSB"
        spc_1.Text = "11MVSC"
        SPC_3.Text = "11MVSC"
        SPC_4.Text = "11MVSC"
    End If

'If Power System 12 is selected
    If cboPS.value = "Power System 12" Then
        BS1.Text = "12S-BS1"
        BS1_1.Text = "12S-BS1"
        GSBA.Text = "12GSBA"
        GSBA_2.Text = "12GSBA"
        GSBA_3.Text = "12GSBA"
        GSBA_4.Text = "12GSBA"
        GSBA_5.Text = "12GSBA"
        GSBB.Text = "12GSBB"
        GSBB_3.Text = "12GSBB"
        GSBB_4.Text = "12GSBB"
        GSBB_5.Text = "12GSBB"
        GSBB_6.Text = "12GSBB"
        GSBC.Text = "12GSBC"
        GSBC_2.Text = "12GSBC"
        GSBC_3.Text = "12GSBC"
        GSBC_4.Text = "12GSBC"
        GSBC_5.Text = "12GSBC"
        SPA.Text = "12MVSA"
        spa_3.Text = "12MVSA"
        SPA_2.Text = "12MVSA"
        SPB.Text = "12MVSB"
        SPB_2.Text = "12MVSB"
        SPB_3.Text = "12MVSB"
        spc_1.Text = "12MVSC"
        SPC_3.Text = "12MVSC"
        SPC_4.Text = "12MVSC"
    End If
'If Power System 13 is selected
    If cboPS.value = "Power System 13" Then
        BS1.Text = "13S-BS1"
        BS1_1.Text = "13S-BS1"
        GSBA.Text = "13GSBA"
        GSBA_2.Text = "13GSBA"
        GSBA_3.Text = "13GSBA"
        GSBA_4.Text = "13GSBA"
        GSBA_5.Text = "13GSBA"
        GSBB.Text = "13GSBB"
        GSBB_3.Text = "13GSBB"
        GSBB_4.Text = "13GSBB"
        GSBB_5.Text = "13GSBB"
        GSBB_6.Text = "13GSBB"
        GSBC.Text = "13GSBC"
        GSBC_2.Text = "13GSBC"
        GSBC_3.Text = "13GSBC"
        GSBC_4.Text = "13GSBC"
        GSBC_5.Text = "13GSBC"
        SPA.Text = "13MVSA"
        spa_3.Text = "13MVSA"
        SPA_2.Text = "13MVSA"
        SPB.Text = "13MVSB"
        SPB_2.Text = "13MVSB"
        SPB_3.Text = "13MVSB"
        spc_1.Text = "13MVSC"
        SPC_3.Text = "13MVSC"
        SPC_4.Text = "13MVSC"
    End If
'If Power System 14 is selected
    If cboPS.value = "Power System 14" Then
        BS1.Text = "14S-BS1"
        BS1_1.Text = "14S-BS1"
        GSBA.Text = "14GSBA"
        GSBA_2.Text = "14GSBA"
        GSBA_3.Text = "14GSBA"
        GSBA_4.Text = "14GSBA"
        GSBA_5.Text = "14GSBA"
        GSBB.Text = "14GSBB"
        GSBB_3.Text = "14GSBB"
        GSBB_4.Text = "14GSBB"
        GSBB_5.Text = "14GSBB"
        GSBB_6.Text = "14GSBB"
        GSBC.Text = "14GSBC"
        GSBC_2.Text = "14GSBC"
        GSBC_3.Text = "14GSBC"
        GSBC_4.Text = "14GSBC"
        GSBC_5.Text = "14GSBC"
        SPA.Text = "14MVSA"
        spa_3.Text = "14MVSA"
        SPA_2.Text = "14MVSA"
        SPB.Text = "14MVSB"
        SPB_2.Text = "14MVSB"
        SPB_3.Text = "14MVSB"
        spc_1.Text = "14MVSC"
        SPC_3.Text = "14MVSC"
        SPC_4.Text = "14MVSC"
    End If
'If LAS 9 Power System 1 is selected
    If cboPS.value = "Power System 1  " Then
        BS1.Text = "15S-BS1"
        BS1_1.Text = "15S-BS1"
        GSBA.Text = "15GSBA"
        GSBA_2.Text = "15GSBA"
        GSBA_3.Text = "15GSBA"
        GSBA_4.Text = "15GSBA"
        GSBA_5.Text = "15GSBA"
        GSBB.Text = "15GSBB"
        GSBB_3.Text = "15GSBB"
        GSBB_4.Text = "15GSBB"
        GSBB_5.Text = "15GSBB"
        GSBB_6.Text = "15GSBB"
        GSBC.Text = "15GSBC"
        GSBC_2.Text = "15GSBC"
        GSBC_3.Text = "15GSBC"
        GSBC_4.Text = "15GSBC"
        GSBC_5.Text = "15GSBC"
        SPA.Text = "15MVSA"
        spa_3.Text = "15MVSA"
        SPA_2.Text = "15MVSA"
        SPB.Text = "15MVSB"
        SPB_2.Text = "15MVSB"
        SPB_3.Text = "15MVSB"
        spc_1.Text = "15MVSC"
        SPC_3.Text = "15MVSC"
        SPC_4.Text = "15MVSC"
    End If
'If LAS 9 Power System 2 is selected
    If cboPS.value = "Power System 2  " Then
        BS1.Text = "16S-BS1"
        BS1_1.Text = "16S-BS1"
        GSBA.Text = "16GSBA"
        GSBA_2.Text = "16GSBA"
        GSBA_3.Text = "16GSBA"
        GSBA_4.Text = "16GSBA"
        GSBA_5.Text = "16GSBA"
        GSBB.Text = "16GSBB"
        GSBB_3.Text = "16GSBB"
        GSBB_4.Text = "16GSBB"
        GSBB_5.Text = "16GSBB"
        GSBB_6.Text = "16GSBB"
        GSBC.Text = "16GSBC"
        GSBC_2.Text = "16GSBC"
        GSBC_3.Text = "16GSBC"
        GSBC_4.Text = "16GSBC"
        GSBC_5.Text = "16GSBC"
        SPA.Text = "16MVSA"
        spa_3.Text = "16MVSA"
        SPA_2.Text = "16MVSA"
        SPB.Text = "16MVSB"
        SPB_2.Text = "16MVSB"
        SPB_3.Text = "16MVSB"
        spc_1.Text = "16MVSC"
        SPC_3.Text = "16MVSC"
        SPC_4.Text = "16MVSC"
    End If
'If LAS 9 Power System 3 is selected
    If cboPS.value = "Power System 3  " Then
        BS1.Text = "17S-BS1"
        BS1_1.Text = "17S-BS1"
        GSBA.Text = "17GSBA"
        GSBA_2.Text = "17GSBA"
        GSBA_3.Text = "17GSBA"
        GSBA_4.Text = "17GSBA"
        GSBA_5.Text = "17GSBA"
        GSBB.Text = "17GSBB"
        GSBB_3.Text = "17GSBB"
        GSBB_4.Text = "17GSBB"
        GSBB_5.Text = "17GSBB"
        GSBB_6.Text = "17GSBB"
        GSBC.Text = "17GSBC"
        GSBC_2.Text = "17GSBC"
        GSBC_3.Text = "17GSBC"
        GSBC_4.Text = "17GSBC"
        GSBC_5.Text = "17GSBC"
        SPA.Text = "17MVSA"
        spa_3.Text = "17MVSA"
        SPA_2.Text = "17MVSA"
        SPB.Text = "17MVSB"
        SPB_2.Text = "17MVSB"
        SPB_3.Text = "17MVSB"
        spc_1.Text = "17MVSC"
        SPC_3.Text = "17MVSC"
        SPC_4.Text = "17MVSC"
    End If
'If LAS 9 Power System 4 is selected
    If cboPS.value = "Power System 4  " Then
        BS1.Text = "18S-BS1"
        BS1_1.Text = "18S-BS1"
        GSBA.Text = "18GSBA"
        GSBA_2.Text = "18GSBA"
        GSBA_3.Text = "18GSBA"
        GSBA_4.Text = "18GSBA"
        GSBA_5.Text = "18GSBA"
        GSBB.Text = "18GSBB"
        GSBB_3.Text = "18GSBB"
        GSBB_4.Text = "18GSBB"
        GSBB_5.Text = "18GSBB"
        GSBB_6.Text = "18GSBB"
        GSBC.Text = "18GSBC"
        GSBC_2.Text = "18GSBC"
        GSBC_3.Text = "18GSBC"
        GSBC_4.Text = "18GSBC"
        GSBC_5.Text = "18GSBC"
        SPA.Text = "18MVSA"
        spa_3.Text = "18MVSA"
        SPA_2.Text = "18MVSA"
        SPB.Text = "18MVSB"
        SPB_2.Text = "18MVSB"
        SPB_3.Text = "18MVSB"
        spc_1.Text = "18MVSC"
        SPC_3.Text = "18MVSC"
        SPC_4.Text = "18MVSC"
        
    End If
'If LAS 10 Power System 1 is selected
    If cboPS.value = " Power System 1" Then
        BS1.Text = "1S-BS1"
        BS1_1.Text = "1S-BS1"
        GSBA.Text = "1GSBA"
        GSBA_2.Text = "1GSBA"
        GSBA_3.Text = "1GSBA"
        GSBA_4.Text = "1GSBA"
        GSBA_5.Text = "1GSBA"
        GSBB.Text = "1GSBB"
        GSBB_3.Text = "1GSBB"
        GSBB_4.Text = "1GSBB"
        GSBB_5.Text = "1GSBB"
        GSBB_6.Text = "1GSBB"
        GSBC.Text = "1GSBC"
        GSBC_2.Text = "1GSBC"
        GSBC_3.Text = "1GSBC"
        GSBC_4.Text = "1GSBC"
        GSBC_5.Text = "1GSBC"
        SPA.Text = "1MVSA"
        spa_3.Text = "1MVSA"
        SPA_2.Text = "1MVSA"
        SPB.Text = "1MVSB"
        SPB_2.Text = "1MVSB"
        SPB_3.Text = "1MVSB"
        spc_1.Text = "1MVSC"
        SPC_3.Text = "1MVSC"
        SPC_4.Text = "1MVSC"
    End If
'If LAS 10 Power System 2 is selected
    If cboPS.value = " Power System 2" Then
        BS1.Text = "2S-BS1"
        BS1_1.Text = "2S-BS1"
        GSBA.Text = "2GSBA"
        GSBA_2.Text = "2GSBA"
        GSBA_3.Text = "2GSBA"
        GSBA_4.Text = "2GSBA"
        GSBA_5.Text = "2GSBA"
        GSBB.Text = "2GSBB"
        GSBB_3.Text = "2GSBB"
        GSBB_4.Text = "2GSBB"
        GSBB_5.Text = "2GSBB"
        GSBB_6.Text = "2GSBB"
        GSBC.Text = "2GSBC"
        GSBC_2.Text = "2GSBC"
        GSBC_3.Text = "2GSBC"
        GSBC_4.Text = "2GSBC"
        GSBC_5.Text = "2GSBC"
        SPA.Text = "2MVSA"
        spa_3.Text = "2MVSA"
        SPA_2.Text = "2MVSA"
        SPB.Text = "2MVSB"
        SPB_2.Text = "2MVSB"
        SPB_3.Text = "2MVSB"
        spc_1.Text = "2MVSC"
        SPC_3.Text = "2MVSC"
        SPC_4.Text = "2MVSC"
    End If
'If LAS 10 Power System 3 is selected
    If cboPS.value = " Power System 3" Then
        BS1.Text = "3S-BS1"
        BS1_1.Text = "3S-BS1"
        GSBA.Text = "3GSBA"
        GSBA_2.Text = "3GSBA"
        GSBA_3.Text = "3GSBA"
        GSBA_4.Text = "3GSBA"
        GSBA_5.Text = "3GSBA"
        GSBB.Text = "3GSBB"
        GSBB_3.Text = "3GSBB"
        GSBB_4.Text = "3GSBB"
        GSBB_5.Text = "3GSBB"
        GSBB_6.Text = "3GSBB"
        GSBC.Text = "3GSBC"
        GSBC_2.Text = "3GSBC"
        GSBC_3.Text = "3GSBC"
        GSBC_4.Text = "3GSBC"
        GSBC_5.Text = "3GSBC"
        SPA.Text = "3MVSA"
        spa_3.Text = "3MVSA"
        SPA_2.Text = "3MVSA"
        SPB.Text = "3MVSB"
        SPB_2.Text = "3MVSB"
        SPB_3.Text = "3MVSB"
        spc_1.Text = "3MVSC"
        SPC_3.Text = "3MVSC"
        SPC_4.Text = "3MVSC"
    End If
'If LAS 11 Power System 1 is selected
    If cboPS.value = "Power System 1" Then
        BS1.Text = "1S-BS1"
        BS1_1.Text = "1S-BS1"
        GSBA.Text = "1GSBA"
        GSBA_2.Text = "1GSBA"
        GSBA_3.Text = "1GSBA"
        GSBA_4.Text = "1GSBA"
        GSBA_5.Text = "1GSBA"
        GSBB.Text = "1GSBB"
        GSBB_3.Text = "1GSBB"
        GSBB_4.Text = "1GSBB"
        GSBB_5.Text = "1GSBB"
        GSBB_6.Text = "1GSBB"
        GSBC.Text = "1GSBC"
        GSBC_2.Text = "1GSBC"
        GSBC_3.Text = "1GSBC"
        GSBC_4.Text = "1GSBC"
        GSBC_5.Text = "1GSBC"
        SPA.Text = "1MVSA"
        spa_3.Text = "1MVSA"
        SPA_2.Text = "1MVSA"
        SPB.Text = "1MVSB"
        SPB_2.Text = "1MVSB"
        SPB_3.Text = "1MVSB"
        spc_1.Text = "1MVSC"
        SPC_3.Text = "1MVSC"
        SPC_4.Text = "1MVSC"
    End If
'If LAS 11 Power System 2 is selected
    If cboPS.value = "Power System 2" Then
        BS1.Text = "2S-BS1"
        BS1_1.Text = "2S-BS1"
        GSBA.Text = "2GSBA"
        GSBA_2.Text = "2GSBA"
        GSBA_3.Text = "2GSBA"
        GSBA_4.Text = "2GSBA"
        GSBA_5.Text = "2GSBA"
        GSBB.Text = "2GSBB"
        GSBB_3.Text = "2GSBB"
        GSBB_4.Text = "2GSBB"
        GSBB_5.Text = "2GSBB"
        GSBB_6.Text = "2GSBB"
        GSBC.Text = "2GSBC"
        GSBC_2.Text = "2GSBC"
        GSBC_3.Text = "2GSBC"
        GSBC_4.Text = "2GSBC"
        GSBC_5.Text = "2GSBC"
        SPA.Text = "2MVSA"
        spa_3.Text = "2MVSA"
        SPA_2.Text = "2MVSA"
        SPB.Text = "2MVSB"
        SPB_2.Text = "2MVSB"
        SPB_3.Text = "2MVSB"
        spc_1.Text = "2MVSC"
        SPC_3.Text = "2MVSC"
        SPC_4.Text = "2MVSC"
        
    End If
'If LAS 11 Power System 3 is selected
    If cboPS.value = "Power System 3" Then
        BS1.Text = "3S-BS1"
        BS1_1.Text = "3S-BS1"
        GSBA.Text = "3GSBA"
        GSBA_2.Text = "3GSBA"
        GSBA_3.Text = "3GSBA"
        GSBA_4.Text = "3GSBA"
        GSBA_5.Text = "3GSBA"
        GSBB.Text = "3GSBB"
        GSBB_3.Text = "3GSBB"
        GSBB_4.Text = "3GSBB"
        GSBB_5.Text = "3GSBB"
        GSBB_6.Text = "3GSBB"
        GSBC.Text = "3GSBC"
        GSBC_2.Text = "3GSBC"
        GSBC_3.Text = "3GSBC"
        GSBC_4.Text = "3GSBC"
        GSBC_5.Text = "3GSBC"
        SPA.Text = "3MVSA"
        spa_3.Text = "3MVSA"
        SPA_2.Text = "3MVSA"
        SPB.Text = "3MVSB"
        SPB_2.Text = "3MVSB"
        SPB_3.Text = "3MVSB"
        spc_1.Text = "3MVSC"
        SPC_3.Text = "3MVSC"
        SPC_4.Text = "3MVSC"
    End If
        
End Sub


Private Sub cboEquipmentID_2_Change()

If cboEquipmentID.value = cboEquipmentID_2.value Then

MsgBox "Please make another selection for the second UPS.", vbOKOnly
 If vbOK = 1 Then
 cboEquipmentID_2.Clear
 cboPS_Change
 End If
End If
 
End Sub

Public Sub cbonumberofups_AfterUpdate()

If cbonumberofups.value = "1" Then
    cbotypeofmaintenance.Clear
    With cbotypeofmaintenance
        .AddItem "1 UPS Annual PM w/o Cal or Depletion"
        .AddItem "1 UPS Annual PM w/ Cal"
        .AddItem "1 UPS Annual PM w/ Cal and Depletion"
        .AddItem "1 UPS Corrective Maintenance"
    End With
    cboEquipmentID_2.Enabled = False
    cboEquipmentID_2.Visible = False
    Label11.Visible = False
ElseIf cbonumberofups.value = "2" Then
    cbotypeofmaintenance.Clear
     With cbotypeofmaintenance
        .AddItem "2 UPS's Annual PM w/o Cal or Depletion"
        .AddItem "2 UPS's Annual PM w/ Cal"
        .AddItem "2 UPS's Annual PM w/ Cal and Depletion"
        .AddItem "2 UPS's Corrective Maintenance"
    End With
    cboEquipmentID_2.Enabled = True
    cboEquipmentID_2.Visible = True
    Label11.Visible = True
End If

End Sub


Public Sub cboSite_Change()

If cboSite.value = "LAS 7" Then
    cboPS.Clear
    With cboPS
    .AddItem "Power System 1 "
    .AddItem "Power System 2 "
    .AddItem "Power System 3 "
    .AddItem "Power System 5"
    .AddItem "Power System 6"
    .AddItem "Power System 7"
    .AddItem "Power System 8"
    End With
End If

If cboSite.value = "LAS 8" Then
    cboPS.Clear
    With cboPS
    .AddItem "Power System 11"
    .AddItem "Power System 12"
    .AddItem "Power System 13"
    .AddItem "Power System 14"
    End With
End If

If cboSite.value = "LAS 9" Then
    cboPS.Clear
    With cboPS
    .AddItem "Power System 1  "
    .AddItem "Power System 2  "
    .AddItem "Power System 3  "
    .AddItem "Power System 4  "
    End With
End If

If cboSite.value = "LAS 10" Then
    cboPS.Clear
    With cboPS
    .AddItem " Power System 1"
    .AddItem " Power System 2"
    .AddItem " Power System 3"
    End With
End If
    
If cboSite.value = "LAS 11" Then
    cboPS.Clear
    With cboPS
    .AddItem "Power System 1"
    .AddItem "Power System 2"
    .AddItem "Power System 3"
    End With
    
End If

End Sub



Public Sub Clear_Click()

    Public ctl As MSForms.Control

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Text = ""
            Case "CheckBox", "OptionButton", "ToggleButton"
                ctl.value = False
            Case "ComboBox", "ListBox"
                ctl.ListIndex = -1
        End Select
    Next ctl
    
cbotypeofmaintenance.Clear
cboEquipmentID.Clear
cboEquipmentID_2.Clear
cboEquipmentID_2.Visible = True
cboEquipmentID_2.Enabled = True
Label11.Visible = True
    
End Sub



Public Sub UserForm_Initialize()

With cbonumberofups
    .AddItem "1"
    .AddItem "2"
End With

With cbocriticalitylevel
    .AddItem "High"
    .AddItem "Low"
End With

With cboSite
    .AddItem "LAS 7"
    .AddItem "LAS 8"
    .AddItem "LAS 9"
    .AddItem "LAS 10"
    .AddItem "LAS 11"
End With


End Sub

Private Sub PDUselection()

If cbonumberofups.value = "1" Then
    
    With ActiveDocument.Tables(6)
    For i = 6 To 9
        ActiveDocument.Tables(6).rows(i).Select
        Selection.rows.Delete
    Next i
        
    End With
End If

End Sub
Public Sub cboPS_Change()

'Populate UPS drop down boxes

If cboPS.value = "Power System 1 " Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = "Power System 1 " Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = "Power System 2 " Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 7A"
    .AddItem "UPS 8A"
    .AddItem "UPS 9B"
    .AddItem "UPS 10B"
    .AddItem "UPS 11C"
    .AddItem "UPS 12C"
    End With
End If

If cboPS.value = "Power System 2 " Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 7A"
    .AddItem "UPS 8A"
    .AddItem "UPS 9B"
    .AddItem "UPS 10B"
    .AddItem "UPS 11C"
    .AddItem "UPS 12C"
    End With
End If

If cboPS.value = "Power System 3 " Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 14A"
    .AddItem "UPS 16B"
    .AddItem "UPS 18C"
    End With
End If

If cboPS.value = "Power System 3 " Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 14A"
    .AddItem "UPS 16B"
    .AddItem "UPS 18C"
    End With
End If

If cboPS.value = "Power System 5" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 25A"
    .AddItem "UPS 26A"
    .AddItem "UPS 27B"
    .AddItem "UPS 28B"
    .AddItem "UPS 29C"
    .AddItem "UPS 30C"
    End With
End If

If cboPS.value = "Power System 5" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 25A"
    .AddItem "UPS 26A"
    .AddItem "UPS 27B"
    .AddItem "UPS 28B"
    .AddItem "UPS 29C"
    .AddItem "UPS 30C"
    End With
End If

If cboPS.value = "Power System 6" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 31A"
    .AddItem "UPS 32A"
    .AddItem "UPS 33B"
    .AddItem "UPS 34B"
    .AddItem "UPS 35C"
    .AddItem "UPS 36C"
    End With
End If

If cboPS.value = "Power System 6" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 31A"
    .AddItem "UPS 32A"
    .AddItem "UPS 33B"
    .AddItem "UPS 34B"
    .AddItem "UPS 35C"
    .AddItem "UPS 36C"
    End With
End If

If cboPS.value = "Power System 7" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 37A"
    .AddItem "UPS 39B"
    .AddItem "UPS 41C"
    End With
End If

If cboPS.value = "Power System 7" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 37A"
    .AddItem "UPS 39B"
    .AddItem "UPS 41C"
    End With
End If

If cboPS.value = "Power System 8" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 43A"
    .AddItem "UPS 44A"
    .AddItem "UPS 45B"
    .AddItem "UPS 46B"
    .AddItem "UPS 47C"
    .AddItem "UPS 48C"
    End With
End If

If cboPS.value = "Power System 8" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 43A"
    .AddItem "UPS 44A"
    .AddItem "UPS 45B"
    .AddItem "UPS 46B"
    .AddItem "UPS 47C"
    .AddItem "UPS 48C"
    End With
End If

If cboPS.value = "Power System 11" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 61A"
    .AddItem "UPS 62A"
    .AddItem "UPS 63B"
    .AddItem "UPS 64B"
    .AddItem "UPS 65C"
    .AddItem "UPS 66C"
    End With
End If

If cboPS.value = "Power System 11" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 61A"
    .AddItem "UPS 62A"
    .AddItem "UPS 63B"
    .AddItem "UPS 64B"
    .AddItem "UPS 65C"
    .AddItem "UPS 66C"
    End With
End If

If cboPS.value = "Power System 12" Then
   cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 67A"
    .AddItem "UPS 68A"
    .AddItem "UPS 69B"
    .AddItem "UPS 70B"
    .AddItem "UPS 71C"
    .AddItem "UPS 72C"
    End With
End If

If cboPS.value = "Power System 12" Then
   cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 67A"
    .AddItem "UPS 68A"
    .AddItem "UPS 69B"
    .AddItem "UPS 70B"
    .AddItem "UPS 71C"
    .AddItem "UPS 72C"
    End With
End If

If cboPS.value = "Power System 13" Then
   cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 73A"
    .AddItem "UPS 74A"
    .AddItem "UPS 75B"
    .AddItem "UPS 76B"
    .AddItem "UPS 77C"
    .AddItem "UPS 78C"
    End With
End If

If cboPS.value = "Power System 13" Then
   cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 73A"
    .AddItem "UPS 74A"
    .AddItem "UPS 75B"
    .AddItem "UPS 76B"
    .AddItem "UPS 77C"
    .AddItem "UPS 78C"
    End With
End If

If cboPS.value = "Power System 14" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 79A"
    .AddItem "UPS 80A"
    .AddItem "UPS 81B"
    .AddItem "UPS 82B"
    .AddItem "UPS 83C"
    .AddItem "UPS 84C"
    End With
End If

If cboPS.value = "Power System 14" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 79A"
    .AddItem "UPS 80A"
    .AddItem "UPS 81B"
    .AddItem "UPS 82B"
    .AddItem "UPS 83C"
    .AddItem "UPS 84C"
    End With
End If

'LAS 9 Power Systems
If cboPS.value = "Power System 1  " Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 85A"
    .AddItem "UPS 86A"
    .AddItem "UPS 87B"
    .AddItem "UPS 88B"
    .AddItem "UPS 89C"
    .AddItem "UPS 90C"
    End With
End If

If cboPS.value = "Power System 1  " Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 85A"
    .AddItem "UPS 86A"
    .AddItem "UPS 87B"
    .AddItem "UPS 88B"
    .AddItem "UPS 89C"
    .AddItem "UPS 90C"
    End With
End If

If cboPS.value = "Power System 2  " Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 91A"
    .AddItem "UPS 92A"
    .AddItem "UPS 94B"
    .AddItem "UPS 95B"
    .AddItem "UPS 96C"
    .AddItem "UPS 97C"
    End With
End If

If cboPS.value = "Power System 2  " Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 91A"
    .AddItem "UPS 92A"
    .AddItem "UPS 94B"
    .AddItem "UPS 95B"
    .AddItem "UPS 96C"
    .AddItem "UPS 97C"
    End With
End If

If cboPS.value = "Power System 3  " Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 98A"
    .AddItem "UPS 99A"
    .AddItem "UPS 100B"
    .AddItem "UPS 101B"
    .AddItem "UPS 102C"
    .AddItem "UPS 103C"
    End With
End If

If cboPS.value = "Power System 3  " Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 98A"
    .AddItem "UPS 99A"
    .AddItem "UPS 100B"
    .AddItem "UPS 101B"
    .AddItem "UPS 102C"
    .AddItem "UPS 103C"
    End With
End If

If cboPS.value = "Power System 4  " Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 104A"
    .AddItem "UPS 105A"
    .AddItem "UPS 106B"
    .AddItem "UPS 107B"
    .AddItem "UPS 108C"
    .AddItem "UPS 109C"
    End With
End If

If cboPS.value = "Power System 4  " Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 104A"
    .AddItem "UPS 105A"
    .AddItem "UPS 106B"
    .AddItem "UPS 107B"
    .AddItem "UPS 108C"
    .AddItem "UPS 109C"
    End With
End If

'LAS 10 Power Systems
If cboPS.value = " Power System 1" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = " Power System 1" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = " Power System 2" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = " Power System 2" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = " Power System 3" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = " Power System 3" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

'LAS 11 Power Systems
If cboPS.value = "Power System 1" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = "Power System 1" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = "Power System 2" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = "Power System 2" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = "Power System 3" Then
    cboEquipmentID.Clear
    With cboEquipmentID
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If

If cboPS.value = "Power System 3" Then
    cboEquipmentID_2.Clear
    With cboEquipmentID_2
    .AddItem "UPS 1A"
    .AddItem "UPS 2A"
    .AddItem "UPS 3B"
    .AddItem "UPS 4B"
    .AddItem "UPS 5C"
    .AddItem "UPS 6C"
    End With
End If
    


End Sub
Public Sub setmbs()
    
    Set MBS = ActiveDocument.Bookmarks("tMBS").Range
    If cboPS.value = "Power System 11" Then
        MBS.Text = "11MBS"
    ElseIf cboPS.value = "Power System 12" Then
        MBS.Text = "12MBS"
    ElseIf cboPS.value = "Power System 13" Then
        MBS.Text = "13MBS"
    ElseIf cboPS.value = "Power System 14" Then
        MBS.Text = "14MBS"
    ElseIf cboPS.value = "Power System 1  " Then
        MBS.Text = "15MBS"
    ElseIf cboPS.value = "Power System 2  " Then
        MBS.Text = "16MBS"
    ElseIf cboPS.value = "Power System 3  " Then
        MBS.Text = "17MBS"
    ElseIf cboPS.value = "Power System 4  " Then
        MBS.Text = "18MBS"
    ElseIf cboPS.value = " Power System 1" Then
        MBS.Text = "1MBS"
    ElseIf cboPS.value = " Power System 2" Then
        MBS.Text = "2MBS"
    ElseIf cboPS.value = " Power System 3" Then
        MBS.Text = "3MBS"
    ElseIf cboPS.value = "Power System 1" Then
        MBS.Text = "1MBS"
    ElseIf cboPS.value = "Power System 2" Then
        MBS.Text = "2MBS"
    ElseIf cboPS.value = "Power System 3" Then
        MBS.Text = "3MBS"
    End If
End Sub
 Public Sub setISX()
 'Disabling and Enabling ISX Statements if its a Red Transfer

Set upscolor = ActiveDocument.Bookmarks("tcolor").Range
Set Title = ActiveDocument.Bookmarks("ttitle").Range
Set MBS_GSB = ActiveDocument.Bookmarks("tMBS_GSB").Range
Set MVS = ActiveDocument.Bookmarks("tMVS").Range
Set ISX = ActiveDocument.Bookmarks("tISX").Range
Set ISX_2 = ActiveDocument.Bookmarks("tISX_2").Range
Set PDU_1 = ActiveDocument.Bookmarks("tPDU_1").Range
Set PDU_2 = ActiveDocument.Bookmarks("tPDU_2").Range
Set PDU_3 = ActiveDocument.Bookmarks("tPDU_3").Range
Set PDU_4 = ActiveDocument.Bookmarks("tPDU_4").Range
Set PDU_5 = ActiveDocument.Bookmarks("tPDU_5").Range
Set PDU_6 = ActiveDocument.Bookmarks("tPDU_6").Range
Set PDU_7 = ActiveDocument.Bookmarks("tPDU_7").Range
Set PDU_8 = ActiveDocument.Bookmarks("tPDU_8").Range
    
    If cboEquipmentID.value = "UPS 61A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "11MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "11MBS-11GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 62A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "11MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "11MBS-11GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 67A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "12MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "12MBS-12GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 68A" Then
           
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "12MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "12MBS-12GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 73A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "13MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "13MBS-13GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 74A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "13MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "13MBS-13GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 79A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "14MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "14MBS-14GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 80A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "14MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "14MBS-14GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 85A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "15MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "15MBS-15GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 86A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "15MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "15MBS-15GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 91A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "16MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "16MBS-16GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 92A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "16MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "16MBS-16GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 97A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "17MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "17MBS-17GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 98A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "17MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "17MBS-17GSBA"

    ElseIf cboEquipmentID.value = "UPS 104A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "18MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "18MBS-18GSBA"
        
    ElseIf cboEquipmentID.value = "UPS 105A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "18MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "18MBS-18GSBA"
        
    'LAS 10 Red UPS's
    
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 1" And cboEquipmentID.value = "UPS 1A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 1 " And cboEquipmentID.value = "UPS 2A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 2 " And cboEquipmentID.value = "UPS 1A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "2MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "2MBS-2GSBA"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 2 " And cboEquipmentID.value = "UPS 2A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "2MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "2MBS-2GSBA"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 3 " And cboEquipmentID.value = "UPS 1A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "3MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "3MBS-3GSBA"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 3 " And cboEquipmentID.value = "UPS 2A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "3MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "3MBS-3GSBA"
        
    'LAS 11 Red UPS's
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 1" And cboEquipmentID.value = "UPS 1A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
    
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 2" And cboEquipmentID.value = "UPS 1A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 3" And cboEquipmentID.value = "UPS 1A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 1" And cboEquipmentID.value = "UPS 2A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 2" And cboEquipmentID.value = "UPS 2A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 3" And cboEquipmentID.value = "UPS 2A" Then
        
        Title.Text = cboSite.value & " Red (A) Building Transfer Script"
        ISX.Text = "Both Red UPS ISX"
        ISX_2.Text = "Both Red UPS ISX"
        MVS.Text = "1MVSA"
        upscolor.Text = "Red"
        MBS_GSB.Text = "1MBS-1GSBA"
        
 End If
    
    'Disabling and Enabling ISX Statements if its a Blue Transfer
    'LAS 8 Blue UPS's
    
    If cboEquipmentID.value = "UPS 63B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "11MVSA"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "11MBS-11GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 64B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "11MVSA"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "11MBS-11GSBB"
    
    ElseIf cboEquipmentID.value = "UPS 69B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "12MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "12MBS-12GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 70B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "12MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "12MBS-12GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 75B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "13MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "13MBS-13GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 76B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "13MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "13MBS-13GSBB"

    ElseIf cboEquipmentID.value = "UPS 81B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "14MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "14MBS-14GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 82B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "14MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "14MBS-14GSBB"
        
    'LAS 9 Blue UPS's
        
    ElseIf cboEquipmentID.value = "UPS 87B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "15MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "15MBS-15GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 88B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "15MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "15MBS-15GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 93B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "16MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "16MBS-16GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 94B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "16MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "16MBS-16GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 99B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "17MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "17MBS-17GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 100B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "17MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "17MBS-17GSBB"
    
    ElseIf cboEquipmentID.value = "UPS 106B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "18MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "18MBS-18GSBB"
        
    ElseIf cboEquipmentID.value = "UPS 107B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "18MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "18MBS-18GSBB"
    
    'LAS 10 Blue UPS's
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 1" And cboEquipmentID.value = "UPS 3B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "1MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "1MBS-1GSBB"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 2" And cboEquipmentID.value = "UPS 3B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "2MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "2MBS-2GSBB"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 3" And cboEquipmentID.value = "UPS 3B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "3MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "3MBS-3GSBB"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 1" And cboEquipmentID.value = "UPS 4B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "1MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "1MBS-1GSBB"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 2" And cboEquipmentID.value = "UPS 4B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "2MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "2MBS-2GSBB"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 3" And cboEquipmentID.value = "UPS 4B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "3MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "3MBS-3GSBB"
        
    'LAS 11 Blue UPS's
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 1" And cboEquipmentID.value = "UPS 3B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "1MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "1MBS-1GSBB"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 2" And cboEquipmentID.value = "UPS 3B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "1MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "1MBS-1GSBB"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 3" And cboEquipmentID.value = "UPS 3B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "1MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "1MBS-1GSBB"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 1" And cboEquipmentID.value = "UPS 4B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "1MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "1MBS-1GSBB"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 2" And cboEquipmentID.value = "UPS 4B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "2MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "2MBS-2GSBB"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 3" And cboEquipmentID.value = "UPS 4B" Then
        
        Title.Text = cboSite.value & " Blue (B) Building Transfer Script"
        ISX.Text = "Both Blue UPS ISX"
        ISX_2.Text = "Both Blue UPS ISX"
        MVS.Text = "3MVSB"
        upscolor.Text = "Blue"
        MBS_GSB.Text = "3MBS-3GSBB"
        
    End If
       
    'Disabling and Enabling ISX Statements if its a Grey Transfer
    'LAS 8 Grey UPS's
    
    If cboEquipmentID.value = "UPS 65C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "11MVSA"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "11MBS-11GSBC"
        
     ElseIf cboEquipmentID.value = "UPS 66C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "11MVSA"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "11MBS-11GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 71C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "12MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "12MBS-12GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 72C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "12MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "12MBS-12GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 77C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "13MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "13MBS-13GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 78C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "13MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "13MBS-13GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 83C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "14MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "14MBS-14GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 84C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "14MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "14MBS-14GSBC"
        
    'LAS 9 Grey UPS's
        
    ElseIf cboEquipmentID.value = "UPS 89C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "15MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "15MBS-15GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 90C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "15MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "15MBS-15GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 95C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "16MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "16MBS-16GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 96C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "16MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "16MBS-16GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 101C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "17MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "17MBS-17GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 102C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "17MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "17MBS-17GSBC"
    
    ElseIf cboEquipmentID.value = "UPS 108C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "18MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "18MBS-18GSBC"
        
    ElseIf cboEquipmentID.value = "UPS 109C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "18MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "18MBS-18GSBC"
        
    'LAS 10 Grey UPS's
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 1" And cboEquipmentID.value = "UPS 5C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 2" And cboEquipmentID.value = "UPS 5C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "2MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "2MBS-2GSBC"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 3" And cboEquipmentID.value = "UPS 5C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "3MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "3MBS-3GSBC"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 1" And cboEquipmentID.value = "UPS 6C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 2" And cboEquipmentID.value = "UPS 6C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "2MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "2MBS-2GSBC"
        
    ElseIf cboSite.value = "LAS 10" And cboPS.value = " Power System 3" And cboEquipmentID.value = "UPS 6C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "3MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "3MBS-3GSBC"
        
    'LAS 11 Grey UPS's
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 1" And cboEquipmentID.value = "UPS 5C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 2" And cboEquipmentID.value = "UPS 5C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 3" And cboEquipmentID.value = "UPS 5C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 1" And cboEquipmentID.value = "UPS 6C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 2" And cboEquipmentID.value = "UPS 6C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
    ElseIf cboSite.value = "LAS 11" And cboPS.value = "Power System 3" And cboEquipmentID.value = "UPS 6C" Then
        
        Title.Text = cboSite.value & " Grey (C) Building Transfer Script"
        ISX.Text = "Both Grey UPS ISX"
        ISX_2.Text = "Both Grey UPS ISX"
        MVS.Text = "1MVSC"
        upscolor.Text = "Grey"
        MBS_GSB.Text = "1MBS-1GSBC"
        
      
    End If
 End Sub
 Private Sub setpduLAS7()
  'LAS 7 PDU List
    Dim PDU_1 As Range
    Dim PDU_2 As Range
    Dim PDU_3 As Range
    Dim PDU_4 As Range
    Dim PDU_5 As Range
    Dim PDU_6 As Range
    Dim PDU_7 As Range
    Dim PDU_8 As Range
    Dim aLAS7PDUS(0 To 138) As String

    
        aLAS7PDUS(0) = "PDU 1A"
        aLAS7PDUS(1) = "PDU 2A"
        aLAS7PDUS(2) = "PDU 3A"
        aLAS7PDUS(3) = "PDU 4A"
        aLAS7PDUS(4) = "PDU 5A"
        aLAS7PDUS(5) = "PDU 6A"
        aLAS7PDUS(6) = "PDU 7A"
        aLAS7PDUS(7) = "PDU 8A"
        aLAS7PDUS(8) = "PDU 9B"
        aLAS7PDUS(9) = "PDU 10B"
        aLAS7PDUS(10) = "PDU 11B"
        aLAS7PDUS(11) = "PDU 12B"
        aLAS7PDUS(12) = "PDU 13B"
        aLAS7PDUS(13) = "PDU 14B"
        aLAS7PDUS(14) = "PDU 15B"
        aLAS7PDUS(15) = "PDU 16B"
        aLAS7PDUS(16) = "PDU 21C"
        aLAS7PDUS(17) = "PDU 22C"
        aLAS7PDUS(18) = "PDU 23C"
        aLAS7PDUS(19) = "PDU 24C"
        aLAS7PDUS(20) = "PDU 17C"
        aLAS7PDUS(21) = "PDU 18C"
        aLAS7PDUS(22) = "PDU 19C"
        aLAS7PDUS(23) = "PDU 20C"
        aLAS7PDUS(24) = "PDU 25A"
        aLAS7PDUS(25) = "PDU 26A"
        aLAS7PDUS(26) = "PDU 27A"
        aLAS7PDUS(27) = "PDU 28A"
        aLAS7PDUS(28) = "PDU 29A"
        aLAS7PDUS(29) = "PDU 30A"
        aLAS7PDUS(30) = "PDU 31A"
        aLAS7PDUS(31) = "PDU 32A"
        aLAS7PDUS(32) = "PDU 33B"
        aLAS7PDUS(33) = "PDU 34B"
        aLAS7PDUS(34) = "PDU 35B"
        aLAS7PDUS(35) = "PDU 36B"
        aLAS7PDUS(36) = "PDU 37B"
        aLAS7PDUS(37) = "PDU 38B"
        aLAS7PDUS(38) = "PDU 39B"
        aLAS7PDUS(39) = "PDU 40B"
        aLAS7PDUS(40) = "PDU 41C"
        aLAS7PDUS(41) = "PDU 42C"
        aLAS7PDUS(42) = "PDU 43C"
        aLAS7PDUS(43) = "PDU 44C"
        aLAS7PDUS(44) = "PDU 45C"
        aLAS7PDUS(45) = "PDU 46C"
        aLAS7PDUS(46) = "PDU 47C"
        aLAS7PDUS(47) = "PDU 48C"
        aLAS7PDUS(48) = "PDU 53A"
        aLAS7PDUS(49) = "PDU 54A"
        aLAS7PDUS(50) = "PDU 55A"
        aLAS7PDUS(51) = "PDU 56A"
        aLAS7PDUS(52) = "PDU 61B"
        aLAS7PDUS(53) = "PDU 62B"
        aLAS7PDUS(54) = "PDU 63B"
        aLAS7PDUS(55) = "PDU 64B"
        aLAS7PDUS(56) = "PDU 69C"
        aLAS7PDUS(57) = "PDU 70C"
        aLAS7PDUS(58) = "PDU 71C"
        aLAS7PDUS(59) = "PDU 72C"
        aLAS7PDUS(60) = "PDU 97A"
        aLAS7PDUS(61) = "PDU 98A"
        aLAS7PDUS(62) = "PDU 99A"
        aLAS7PDUS(63) = "PDU 100A"
        aLAS7PDUS(64) = "PDU 101A"
        aLAS7PDUS(65) = "PDU 102A"
        aLAS7PDUS(66) = "PDU 103A"
        aLAS7PDUS(67) = "PDU 104A"
        aLAS7PDUS(68) = "PDU 105B"
        aLAS7PDUS(69) = "PDU 106B"
        aLAS7PDUS(70) = "PDU 107B"
        aLAS7PDUS(71) = "PDU 108B"
        aLAS7PDUS(72) = "PDU 109B"
        aLAS7PDUS(73) = "PDU 110B"
        aLAS7PDUS(74) = "PDU 111B"
        aLAS7PDUS(75) = "PDU 112B"
        aLAS7PDUS(76) = "PDU 113C"
        aLAS7PDUS(77) = "PDU 114C"
        aLAS7PDUS(78) = "PDU 115C"
        aLAS7PDUS(79) = "PDU 116C"
        aLAS7PDUS(80) = "PDU 117C"
        aLAS7PDUS(81) = "PDU 118C"
        aLAS7PDUS(82) = "PDU 119C"
        aLAS7PDUS(83) = "PDU 120C"
        aLAS7PDUS(84) = "PDU 121A"
        aLAS7PDUS(85) = "PDU 122A"
        aLAS7PDUS(86) = "PDU 123A"
        aLAS7PDUS(87) = "PDU 124A"
        aLAS7PDUS(88) = "PDU 125A"
        aLAS7PDUS(89) = "PDU 126A"
        aLAS7PDUS(90) = "PDU 127A"
        aLAS7PDUS(91) = "PDU 128A"
        aLAS7PDUS(92) = "PDU 129B"
        aLAS7PDUS(93) = "PDU 130B"
        aLAS7PDUS(94) = "PDU 131B"
        aLAS7PDUS(95) = "PDU 132B"
        aLAS7PDUS(96) = "PDU 133B"
        aLAS7PDUS(97) = "PDU 134B"
        aLAS7PDUS(98) = "PDU 135B"
        aLAS7PDUS(99) = "PDU 136B"
        aLAS7PDUS(100) = "PDU 137C"
        aLAS7PDUS(101) = "PDU 138C"
        aLAS7PDUS(102) = "PDU 139C"
        aLAS7PDUS(103) = "PDU 140C"
        aLAS7PDUS(104) = "PDU 141C"
        aLAS7PDUS(105) = "PDU 142C"
        aLAS7PDUS(106) = "PDU 143C"
        aLAS7PDUS(107) = "PDU 144C"
        aLAS7PDUS(108) = "PDU 145A"
        aLAS7PDUS(109) = "PDU 146A"
        aLAS7PDUS(110) = "PDU 153B"
        aLAS7PDUS(102) = "PDU 154B"
        aLAS7PDUS(103) = "PDU 161C"
        aLAS7PDUS(104) = "PDU 162C"
        aLAS7PDUS(105) = "PDU 169A"
        aLAS7PDUS(106) = "PDU 170A"
        aLAS7PDUS(107) = "PDU 171A"
        aLAS7PDUS(108) = "PDU 172A"
        aLAS7PDUS(109) = "PDU 173A"
        aLAS7PDUS(110) = "PDU 174A"
        aLAS7PDUS(111) = "PDU 175A"
        aLAS7PDUS(112) = "PDU 176A"
        aLAS7PDUS(113) = "PDU 177B"
        aLAS7PDUS(114) = "PDU 178B"
        aLAS7PDUS(115) = "PDU 179B"
        aLAS7PDUS(116) = "PDU 180B"
        aLAS7PDUS(117) = "PDU 181B"
        aLAS7PDUS(118) = "PDU 182B"
        aLAS7PDUS(119) = "PDU 183B"
        aLAS7PDUS(120) = "PDU 184B"
        aLAS7PDUS(121) = "PDU 185C"
        aLAS7PDUS(122) = "PDU 186C"
        aLAS7PDUS(123) = "PDU 187C"
        aLAS7PDUS(124) = "PDU 188C"
        aLAS7PDUS(125) = "PDU 189C"
        aLAS7PDUS(126) = "PDU 190C"
        aLAS7PDUS(127) = "PDU 191C"
        aLAS7PDUS(128) = "PDU 192C"
        
            Set PDU_1 = ActiveDocument.Bookmarks("tPDU_1").Range
            Set PDU_2 = ActiveDocument.Bookmarks("tPDU_2").Range
            Set PDU_3 = ActiveDocument.Bookmarks("tPDU_3").Range
            Set PDU_4 = ActiveDocument.Bookmarks("tPDU_4").Range
            Set PDU_5 = ActiveDocument.Bookmarks("tPDU_5").Range
            Set PDU_6 = ActiveDocument.Bookmarks("tPDU_6").Range
            Set PDU_7 = ActiveDocument.Bookmarks("tPDU_7").Range
            Set PDU_8 = ActiveDocument.Bookmarks("tPDU_8").Range

        'Setting the PDU's for any of the Power Systems Selected for NAP8
        

                If cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 1A" And cboEquipmentID_2.value = _
                "UPS 2A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(0)
                PDU_2.Text = aLAS7PDUS(1)
                PDU_3.Text = aLAS7PDUS(2)
                PDU_4.Text = aLAS7PDUS(3)
                PDU_5.Text = aLAS7PDUS(4)
                PDU_6.Text = aLAS7PDUS(5)
                PDU_7.Text = aLAS7PDUS(6)
                PDU_8.Text = aLAS7PDUS(7)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 1A" And _
                (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(0)
                PDU_2.Text = aLAS7PDUS(1)
                PDU_3.Text = aLAS7PDUS(2)
                PDU_4.Text = aLAS7PDUS(3)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 2A" And _
                (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(4)
                PDU_2.Text = aLAS7PDUS(5)
                PDU_3.Text = aLAS7PDUS(6)
                PDU_4.Text = aLAS7PDUS(7)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 7A" And cboEquipmentID_2.value = _
                "UPS 8A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(24)
                PDU_2.Text = aLAS7PDUS(25)
                PDU_3.Text = aLAS7PDUS(26)
                PDU_4.Text = aLAS7PDUS(27)
                PDU_5.Text = aLAS7PDUS(28)
                PDU_6.Text = aLAS7PDUS(29)
                PDU_7.Text = aLAS7PDUS(30)
                PDU_8.Text = aLAS7PDUS(31)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 7A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(24)
                PDU_2.Text = aLAS7PDUS(25)
                PDU_3.Text = aLAS7PDUS(26)
                PDU_4.Text = aLAS7PDUS(27)

                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 8A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(28)
                PDU_2.Text = aLAS7PDUS(29)
                PDU_3.Text = aLAS7PDUS(30)
                PDU_4.Text = aLAS7PDUS(31)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 14A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(48)
                PDU_2.Text = aLAS7PDUS(49)
                PDU_3.Text = aLAS7PDUS(50)
                PDU_4.Text = aLAS7PDUS(51)

                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 25A" And cboEquipmentID_2.value = _
                "UPS 26A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(60)
                PDU_2.Text = aLAS7PDUS(61)
                PDU_3.Text = aLAS7PDUS(62)
                PDU_4.Text = aLAS7PDUS(63)
                PDU_5.Text = aLAS7PDUS(64)
                PDU_6.Text = aLAS7PDUS(65)
                PDU_7.Text = aLAS7PDUS(66)
                PDU_8.Text = aLAS7PDUS(67)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 25A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(60)
                PDU_2.Text = aLAS7PDUS(61)
                PDU_3.Text = aLAS7PDUS(62)
                PDU_4.Text = aLAS7PDUS(63)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 26A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(64)
                PDU_2.Text = aLAS7PDUS(65)
                PDU_3.Text = aLAS7PDUS(66)
                PDU_4.Text = aLAS7PDUS(67)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 31A" And cboEquipmentID_2.value = _
                "UPS 32A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(84)
                PDU_2.Text = aLAS7PDUS(85)
                PDU_3.Text = aLAS7PDUS(86)
                PDU_4.Text = aLAS7PDUS(87)
                PDU_5.Text = aLAS7PDUS(88)
                PDU_6.Text = aLAS7PDUS(89)
                PDU_7.Text = aLAS7PDUS(90)
                PDU_8.Text = aLAS7PDUS(91)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 31A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(84)
                PDU_2.Text = aLAS7PDUS(85)
                PDU_3.Text = aLAS7PDUS(86)
                PDU_4.Text = aLAS7PDUS(87)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 32A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(88)
                PDU_2.Text = aLAS7PDUS(89)
                PDU_3.Text = aLAS7PDUS(90)
                PDU_4.Text = aLAS7PDUS(91)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 37A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(108)
                PDU_2.Text = aLAS7PDUS(109)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 43A" And cboEquipmentID_2.value = _
                "UPS 44A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(105)
                PDU_2.Text = aLAS7PDUS(106)
                PDU_3.Text = aLAS7PDUS(107)
                PDU_4.Text = aLAS7PDUS(108)
                PDU_5.Text = aLAS7PDUS(109)
                PDU_6.Text = aLAS7PDUS(110)
                PDU_7.Text = aLAS7PDUS(111)
                PDU_8.Text = aLAS7PDUS(112)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 43A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(105)
                PDU_2.Text = aLAS7PDUS(106)
                PDU_3.Text = aLAS7PDUS(107)
                PDU_4.Text = aLAS7PDUS(108)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 44A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(109)
                PDU_2.Text = aLAS7PDUS(110)
                PDU_3.Text = aLAS7PDUS(111)
                PDU_4.Text = aLAS7PDUS(112)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 3B" And cboEquipmentID_2.value = _
                "UPS 4B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(8)
                PDU_2.Text = aLAS7PDUS(9)
                PDU_3.Text = aLAS7PDUS(10)
                PDU_4.Text = aLAS7PDUS(11)
                PDU_5.Text = aLAS7PDUS(12)
                PDU_6.Text = aLAS7PDUS(13)
                PDU_7.Text = aLAS7PDUS(14)
                PDU_8.Text = aLAS7PDUS(15)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 3B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(8)
                PDU_2.Text = aLAS7PDUS(9)
                PDU_3.Text = aLAS7PDUS(10)
                PDU_4.Text = aLAS7PDUS(11)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 4B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(12)
                PDU_2.Text = aLAS7PDUS(13)
                PDU_3.Text = aLAS7PDUS(14)
                PDU_4.Text = aLAS7PDUS(15)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 9B" And cboEquipmentID_2.value = _
                "UPS 10B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(32)
                PDU_2.Text = aLAS7PDUS(33)
                PDU_3.Text = aLAS7PDUS(34)
                PDU_4.Text = aLAS7PDUS(35)
                PDU_5.Text = aLAS7PDUS(36)
                PDU_6.Text = aLAS7PDUS(37)
                PDU_7.Text = aLAS7PDUS(38)
                PDU_8.Text = aLAS7PDUS(39)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 9B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(32)
                PDU_2.Text = aLAS7PDUS(33)
                PDU_3.Text = aLAS7PDUS(34)
                PDU_4.Text = aLAS7PDUS(35)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 10B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(36)
                PDU_2.Text = aLAS7PDUS(37)
                PDU_3.Text = aLAS7PDUS(38)
                PDU_4.Text = aLAS7PDUS(39)
                           
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 16B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(52)
                PDU_2.Text = aLAS7PDUS(53)
                PDU_3.Text = aLAS7PDUS(54)
                PDU_4.Text = aLAS7PDUS(55)
                                              
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 27B" And cboEquipmentID_2.value = _
                "UPS 28B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(68)
                PDU_2.Text = aLAS7PDUS(69)
                PDU_3.Text = aLAS7PDUS(70)
                PDU_4.Text = aLAS7PDUS(71)
                PDU_5.Text = aLAS7PDUS(72)
                PDU_6.Text = aLAS7PDUS(73)
                PDU_7.Text = aLAS7PDUS(74)
                PDU_8.Text = aLAS7PDUS(75)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 27B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(68)
                PDU_2.Text = aLAS7PDUS(69)
                PDU_3.Text = aLAS7PDUS(70)
                PDU_4.Text = aLAS7PDUS(71)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 28B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(72)
                PDU_2.Text = aLAS7PDUS(73)
                PDU_3.Text = aLAS7PDUS(74)
                PDU_4.Text = aLAS7PDUS(75)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 33B" And cboEquipmentID_2.value = _
                "UPS 34B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(92)
                PDU_2.Text = aLAS7PDUS(93)
                PDU_3.Text = aLAS7PDUS(94)
                PDU_4.Text = aLAS7PDUS(95)
                PDU_5.Text = aLAS7PDUS(96)
                PDU_6.Text = aLAS7PDUS(97)
                PDU_7.Text = aLAS7PDUS(98)
                PDU_8.Text = aLAS7PDUS(99)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 33B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(92)
                PDU_2.Text = aLAS7PDUS(93)
                PDU_3.Text = aLAS7PDUS(94)
                PDU_4.Text = aLAS7PDUS(95)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 34B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(96)
                PDU_2.Text = aLAS7PDUS(97)
                PDU_3.Text = aLAS7PDUS(98)
                PDU_4.Text = aLAS7PDUS(99)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 39B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(110)
                PDU_2.Text = aLAS7PDUS(111)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 45B" And cboEquipmentID_2.value = _
                "UPS 46B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(113)
                PDU_2.Text = aLAS7PDUS(114)
                PDU_3.Text = aLAS7PDUS(115)
                PDU_4.Text = aLAS7PDUS(116)
                PDU_5.Text = aLAS7PDUS(117)
                PDU_6.Text = aLAS7PDUS(118)
                PDU_7.Text = aLAS7PDUS(119)
                PDU_8.Text = aLAS7PDUS(120)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 45B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(113)
                PDU_2.Text = aLAS7PDUS(114)
                PDU_3.Text = aLAS7PDUS(115)
                PDU_4.Text = aLAS7PDUS(116)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 46B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(117)
                PDU_2.Text = aLAS7PDUS(118)
                PDU_3.Text = aLAS7PDUS(119)
                PDU_4.Text = aLAS7PDUS(120)
                
                 ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 5C" And cboEquipmentID_2.value = _
                "UPS 6C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(16)
                PDU_2.Text = aLAS7PDUS(17)
                PDU_3.Text = aLAS7PDUS(18)
                PDU_4.Text = aLAS7PDUS(19)
                PDU_5.Text = aLAS7PDUS(20)
                PDU_6.Text = aLAS7PDUS(21)
                PDU_7.Text = aLAS7PDUS(22)
                PDU_8.Text = aLAS7PDUS(23)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 5C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(16)
                PDU_2.Text = aLAS7PDUS(17)
                PDU_3.Text = aLAS7PDUS(18)
                PDU_4.Text = aLAS7PDUS(19)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 6C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(20)
                PDU_2.Text = aLAS7PDUS(21)
                PDU_3.Text = aLAS7PDUS(22)
                PDU_4.Text = aLAS7PDUS(23)
                
                 ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 11C" And cboEquipmentID_2.value = _
                "UPS 12C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(40)
                PDU_2.Text = aLAS7PDUS(41)
                PDU_3.Text = aLAS7PDUS(42)
                PDU_4.Text = aLAS7PDUS(43)
                PDU_5.Text = aLAS7PDUS(44)
                PDU_6.Text = aLAS7PDUS(45)
                PDU_7.Text = aLAS7PDUS(46)
                PDU_8.Text = aLAS7PDUS(47)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 11C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(40)
                PDU_2.Text = aLAS7PDUS(41)
                PDU_3.Text = aLAS7PDUS(42)
                PDU_4.Text = aLAS7PDUS(43)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 12C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(44)
                PDU_2.Text = aLAS7PDUS(45)
                PDU_3.Text = aLAS7PDUS(46)
                PDU_4.Text = aLAS7PDUS(47)
                                        
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 18C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(56)
                PDU_2.Text = aLAS7PDUS(57)
                PDU_3.Text = aLAS7PDUS(58)
                PDU_4.Text = aLAS7PDUS(59)
                
                 ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 29C" And cboEquipmentID_2.value = _
                "UPS 30C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(76)
                PDU_2.Text = aLAS7PDUS(77)
                PDU_3.Text = aLAS7PDUS(78)
                PDU_4.Text = aLAS7PDUS(79)
                PDU_5.Text = aLAS7PDUS(80)
                PDU_6.Text = aLAS7PDUS(81)
                PDU_7.Text = aLAS7PDUS(82)
                PDU_8.Text = aLAS7PDUS(83)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 29C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(76)
                PDU_2.Text = aLAS7PDUS(77)
                PDU_3.Text = aLAS7PDUS(78)
                PDU_4.Text = aLAS7PDUS(79)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 30C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(80)
                PDU_2.Text = aLAS7PDUS(81)
                PDU_3.Text = aLAS7PDUS(82)
                PDU_4.Text = aLAS7PDUS(83)
                
                 ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 35C" And cboEquipmentID_2.value = _
                "UPS 36C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(100)
                PDU_2.Text = aLAS7PDUS(101)
                PDU_3.Text = aLAS7PDUS(102)
                PDU_4.Text = aLAS7PDUS(103)
                PDU_5.Text = aLAS7PDUS(104)
                PDU_6.Text = aLAS7PDUS(105)
                PDU_7.Text = aLAS7PDUS(106)
                PDU_8.Text = aLAS7PDUS(107)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 35C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(100)
                PDU_2.Text = aLAS7PDUS(101)
                PDU_3.Text = aLAS7PDUS(102)
                PDU_4.Text = aLAS7PDUS(103)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 36C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(104)
                PDU_2.Text = aLAS7PDUS(105)
                PDU_3.Text = aLAS7PDUS(106)
                PDU_4.Text = aLAS7PDUS(107)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 41C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(103)
                PDU_2.Text = aLAS7PDUS(104)

                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 47C" And cboEquipmentID_2.value = _
                "UPS 48C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(121)
                PDU_2.Text = aLAS7PDUS(122)
                PDU_3.Text = aLAS7PDUS(123)
                PDU_4.Text = aLAS7PDUS(124)
                PDU_5.Text = aLAS7PDUS(125)
                PDU_6.Text = aLAS7PDUS(126)
                PDU_7.Text = aLAS7PDUS(127)
                PDU_8.Text = aLAS7PDUS(128)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 47C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(121)
                PDU_2.Text = aLAS7PDUS(122)
                PDU_3.Text = aLAS7PDUS(123)
                PDU_4.Text = aLAS7PDUS(124)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 48C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS7PDUS(125)
                PDU_2.Text = aLAS7PDUS(126)
                PDU_3.Text = aLAS7PDUS(127)
                PDU_4.Text = aLAS7PDUS(128)
            End If


 End Sub
 Private Sub setpduLAS8()
  'LAS 8 PDU List
    Dim PDU_1 As Range
    Dim PDU_2 As Range
    Dim PDU_3 As Range
    Dim PDU_4 As Range
    Dim PDU_5 As Range
    Dim PDU_6 As Range
    Dim PDU_7 As Range
    Dim PDU_8 As Range
    Dim aLAS8PDUS(0 To 110) As String

    
        aLAS8PDUS(0) = "PDU 241A"
        aLAS8PDUS(1) = "PDU 242A"
        aLAS8PDUS(2) = "PDU 243A"
        aLAS8PDUS(3) = "PDU 244A"
        aLAS8PDUS(4) = "PDU 245A"
        aLAS8PDUS(5) = "PDU 246A"
        aLAS8PDUS(6) = "PDU 247A"
        aLAS8PDUS(7) = "PDU 248A"
        aLAS8PDUS(8) = "PDU 249B"
        aLAS8PDUS(9) = "PDU 250B"
        aLAS8PDUS(10) = "PDU 251B"
        aLAS8PDUS(11) = "PDU 252B"
        aLAS8PDUS(12) = "PDU 253B"
        aLAS8PDUS(13) = "PDU 254B"
        aLAS8PDUS(14) = "PDU 255B"
        aLAS8PDUS(15) = "PDU 256B"
        aLAS8PDUS(16) = "PDU 257C"
        aLAS8PDUS(17) = "PDU 258C"
        aLAS8PDUS(18) = "PDU 259C"
        aLAS8PDUS(19) = "PDU 260C"
        aLAS8PDUS(20) = "PDU 261C"
        aLAS8PDUS(21) = "PDU 262C"
        aLAS8PDUS(22) = "PDU 264C"
        aLAS8PDUS(23) = "PDU 265A"
        aLAS8PDUS(24) = "PDU 266A"
        aLAS8PDUS(25) = "PDU 267A"
        aLAS8PDUS(26) = "PDU 268A"
        aLAS8PDUS(27) = "PDU 269A"
        aLAS8PDUS(28) = "PDU 270A"
        aLAS8PDUS(29) = "PDU 271A"
        aLAS8PDUS(30) = "PDU 272A"
        aLAS8PDUS(31) = "PDU 289A"
        aLAS8PDUS(32) = "PDU 290A"
        aLAS8PDUS(33) = "PDU 291A"
        aLAS8PDUS(34) = "PDU 292A"
        aLAS8PDUS(35) = "PDU 293A"
        aLAS8PDUS(36) = "PDU 294A"
        aLAS8PDUS(37) = "PDU 295A"
        aLAS8PDUS(38) = "PDU 296A"
        aLAS8PDUS(39) = "PDU 313A"
        aLAS8PDUS(40) = "PDU 314A"
        aLAS8PDUS(41) = "PDU 315A"
        aLAS8PDUS(42) = "PDU 316A"
        aLAS8PDUS(43) = "PDU 317A"
        aLAS8PDUS(44) = "PDU 318A"
        aLAS8PDUS(45) = "PDU 319A"
        aLAS8PDUS(46) = "PDU 320A"
        aLAS8PDUS(47) = "PDU 249B"
        aLAS8PDUS(48) = "PDU 250B"
        aLAS8PDUS(49) = "PDU 251B"
        aLAS8PDUS(50) = "PDU 252B"
        aLAS8PDUS(51) = "PDU 253B"
        aLAS8PDUS(52) = "PDU 254B"
        aLAS8PDUS(53) = "PDU 255B"
        aLAS8PDUS(54) = "PDU 256B"
        aLAS8PDUS(55) = "PDU 273B"
        aLAS8PDUS(56) = "PDU 274B"
        aLAS8PDUS(57) = "PDU 275B"
        aLAS8PDUS(58) = "PDU 276B"
        aLAS8PDUS(59) = "PDU 277B"
        aLAS8PDUS(60) = "PDU 278B"
        aLAS8PDUS(61) = "PDU 279B"
        aLAS8PDUS(62) = "PDU 280B"
        aLAS8PDUS(63) = "PDU 297B"
        aLAS8PDUS(64) = "PDU 298B"
        aLAS8PDUS(65) = "PDU 299B"
        aLAS8PDUS(66) = "PDU 300B"
        aLAS8PDUS(67) = "PDU 301B"
        aLAS8PDUS(68) = "PDU 302B"
        aLAS8PDUS(69) = "PDU 303B"
        aLAS8PDUS(70) = "PDU 304B"
        aLAS8PDUS(71) = "PDU 321B"
        aLAS8PDUS(72) = "PDU 322B"
        aLAS8PDUS(73) = "PDU 323B"
        aLAS8PDUS(74) = "PDU 324B"
        aLAS8PDUS(75) = "PDU 325B"
        aLAS8PDUS(76) = "PDU 326B"
        aLAS8PDUS(77) = "PDU 327B"
        aLAS8PDUS(78) = "PDU 328B"
        aLAS8PDUS(79) = "PDU 257C"
        aLAS8PDUS(80) = "PDU 258C"
        aLAS8PDUS(81) = "PDU 259C"
        aLAS8PDUS(82) = "PDU 260C"
        aLAS8PDUS(83) = "PDU 261C"
        aLAS8PDUS(84) = "PDU 262C"
        aLAS8PDUS(85) = "PDU 263C"
        aLAS8PDUS(86) = "PDU 264C"
        aLAS8PDUS(87) = "PDU 281C"
        aLAS8PDUS(88) = "PDU 282C"
        aLAS8PDUS(89) = "PDU 283C"
        aLAS8PDUS(90) = "PDU 284C"
        aLAS8PDUS(91) = "PDU 285C"
        aLAS8PDUS(92) = "PDU 286C"
        aLAS8PDUS(93) = "PDU 287C"
        aLAS8PDUS(94) = "PDU 288C"
        aLAS8PDUS(95) = "PDU 305C"
        aLAS8PDUS(96) = "PDU 306C"
        aLAS8PDUS(97) = "PDU 307C"
        aLAS8PDUS(98) = "PDU 308C"
        aLAS8PDUS(99) = "PDU 309C"
        aLAS8PDUS(100) = "PDU 310C"
        aLAS8PDUS(101) = "PDU 311C"
        aLAS8PDUS(102) = "PDU 312C"
        aLAS8PDUS(103) = "PDU 329C"
        aLAS8PDUS(104) = "PDU 330C"
        aLAS8PDUS(105) = "PDU 331C"
        aLAS8PDUS(106) = "PDU 332C"
        aLAS8PDUS(107) = "PDU 333C"
        aLAS8PDUS(108) = "PDU 334C"
        aLAS8PDUS(109) = "PDU 335C"
        aLAS8PDUS(110) = "PDU 336C"
        
            Set PDU_1 = ActiveDocument.Bookmarks("tPDU_1").Range
            Set PDU_2 = ActiveDocument.Bookmarks("tPDU_2").Range
            Set PDU_3 = ActiveDocument.Bookmarks("tPDU_3").Range
            Set PDU_4 = ActiveDocument.Bookmarks("tPDU_4").Range
            Set PDU_5 = ActiveDocument.Bookmarks("tPDU_5").Range
            Set PDU_6 = ActiveDocument.Bookmarks("tPDU_6").Range
            Set PDU_7 = ActiveDocument.Bookmarks("tPDU_7").Range
            Set PDU_8 = ActiveDocument.Bookmarks("tPDU_8").Range

        'Setting the PDU's for any of the Power Systems Selected for NAP8
        

                If cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 61A" And cboEquipmentID_2.value = _
                "UPS 62A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(0)
                PDU_2.Text = aLAS8PDUS(1)
                PDU_3.Text = aLAS8PDUS(2)
                PDU_4.Text = aLAS8PDUS(3)
                PDU_5.Text = aLAS8PDUS(4)
                PDU_6.Text = aLAS8PDUS(5)
                PDU_7.Text = aLAS8PDUS(6)
                PDU_8.Text = aLAS8PDUS(7)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 61A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(0)
                PDU_2.Text = aLAS8PDUS(1)
                PDU_3.Text = aLAS8PDUS(2)
                PDU_4.Text = aLAS8PDUS(3)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 62A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(4)
                PDU_2.Text = aLAS8PDUS(5)
                PDU_3.Text = aLAS8PDUS(6)
                PDU_4.Text = aLAS8PDUS(7)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 67A" And cboEquipmentID_2.value = _
                "UPS 68A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(23)
                PDU_2.Text = aLAS8PDUS(24)
                PDU_3.Text = aLAS8PDUS(25)
                PDU_4.Text = aLAS8PDUS(26)
                PDU_5.Text = aLAS8PDUS(27)
                PDU_6.Text = aLAS8PDUS(28)
                PDU_7.Text = aLAS8PDUS(29)
                PDU_8.Text = aLAS8PDUS(30)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 67A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(23)
                PDU_2.Text = aLAS8PDUS(24)
                PDU_3.Text = aLAS8PDUS(25)
                PDU_4.Text = aLAS8PDUS(26)

                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 68A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(27)
                PDU_2.Text = aLAS8PDUS(28)
                PDU_3.Text = aLAS8PDUS(29)
                PDU_4.Text = aLAS8PDUS(30)

                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 73A" And cboEquipmentID_2.value = _
                "UPS 74A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(31)
                PDU_2.Text = aLAS8PDUS(32)
                PDU_3.Text = aLAS8PDUS(33)
                PDU_4.Text = aLAS8PDUS(34)
                PDU_5.Text = aLAS8PDUS(35)
                PDU_6.Text = aLAS8PDUS(36)
                PDU_7.Text = aLAS8PDUS(37)
                PDU_8.Text = aLAS8PDUS(38)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 73A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(31)
                PDU_2.Text = aLAS8PDUS(32)
                PDU_3.Text = aLAS8PDUS(33)
                PDU_4.Text = aLAS8PDUS(34)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 74A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(35)
                PDU_2.Text = aLAS8PDUS(36)
                PDU_3.Text = aLAS8PDUS(37)
                PDU_4.Text = aLAS8PDUS(38)

                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 79A" And cboEquipmentID_2.value = _
                "UPS 80A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(39)
                PDU_2.Text = aLAS8PDUS(40)
                PDU_3.Text = aLAS8PDUS(41)
                PDU_4.Text = aLAS8PDUS(42)
                PDU_5.Text = aLAS8PDUS(43)
                PDU_6.Text = aLAS8PDUS(44)
                PDU_7.Text = aLAS8PDUS(45)
                PDU_8.Text = aLAS8PDUS(46)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 79A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(39)
                PDU_2.Text = aLAS8PDUS(40)
                PDU_3.Text = aLAS8PDUS(41)
                PDU_4.Text = aLAS8PDUS(42)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 80A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(43)
                PDU_2.Text = aLAS8PDUS(44)
                PDU_3.Text = aLAS8PDUS(45)
                PDU_4.Text = aLAS8PDUS(46)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 63B" And cboEquipmentID_2.value = _
                "UPS 64B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(8)
                PDU_2.Text = aLAS8PDUS(9)
                PDU_3.Text = aLAS8PDUS(10)
                PDU_4.Text = aLAS8PDUS(11)
                PDU_5.Text = aLAS8PDUS(12)
                PDU_6.Text = aLAS8PDUS(13)
                PDU_7.Text = aLAS8PDUS(14)
                PDU_8.Text = aLAS8PDUS(15)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 63B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(8)
                PDU_2.Text = aLAS8PDUS(9)
                PDU_3.Text = aLAS8PDUS(10)
                PDU_4.Text = aLAS8PDUS(11)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 64B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(12)
                PDU_2.Text = aLAS8PDUS(13)
                PDU_3.Text = aLAS8PDUS(14)
                PDU_4.Text = aLAS8PDUS(15)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 69B" And cboEquipmentID_2.value = _
                "UPS 70B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(55)
                PDU_2.Text = aLAS8PDUS(56)
                PDU_3.Text = aLAS8PDUS(57)
                PDU_4.Text = aLAS8PDUS(58)
                PDU_5.Text = aLAS8PDUS(59)
                PDU_6.Text = aLAS8PDUS(60)
                PDU_7.Text = aLAS8PDUS(61)
                PDU_8.Text = aLAS8PDUS(62)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 69B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(55)
                PDU_2.Text = aLAS8PDUS(56)
                PDU_3.Text = aLAS8PDUS(57)
                PDU_4.Text = aLAS8PDUS(58)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 70B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(59)
                PDU_2.Text = aLAS8PDUS(60)
                PDU_3.Text = aLAS8PDUS(61)
                PDU_4.Text = aLAS8PDUS(62)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 75B" And cboEquipmentID_2.value = _
                "UPS 76B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(63)
                PDU_2.Text = aLAS8PDUS(64)
                PDU_3.Text = aLAS8PDUS(65)
                PDU_4.Text = aLAS8PDUS(66)
                PDU_5.Text = aLAS8PDUS(67)
                PDU_6.Text = aLAS8PDUS(68)
                PDU_7.Text = aLAS8PDUS(69)
                PDU_8.Text = aLAS8PDUS(70)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 75B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(63)
                PDU_2.Text = aLAS8PDUS(64)
                PDU_3.Text = aLAS8PDUS(65)
                PDU_4.Text = aLAS8PDUS(66)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 76B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(67)
                PDU_2.Text = aLAS8PDUS(68)
                PDU_3.Text = aLAS8PDUS(69)
                PDU_4.Text = aLAS8PDUS(70)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 81B" And cboEquipmentID_2.value = _
                "UPS 82B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(71)
                PDU_2.Text = aLAS8PDUS(72)
                PDU_3.Text = aLAS8PDUS(73)
                PDU_4.Text = aLAS8PDUS(74)
                PDU_5.Text = aLAS8PDUS(75)
                PDU_6.Text = aLAS8PDUS(76)
                PDU_7.Text = aLAS8PDUS(77)
                PDU_8.Text = aLAS8PDUS(78)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 81B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(71)
                PDU_2.Text = aLAS8PDUS(72)
                PDU_3.Text = aLAS8PDUS(73)
                PDU_4.Text = aLAS8PDUS(74)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 82B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(75)
                PDU_2.Text = aLAS8PDUS(76)
                PDU_3.Text = aLAS8PDUS(77)
                PDU_4.Text = aLAS8PDUS(78)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 65C" And cboEquipmentID_2.value = _
                "UPS 66C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(16)
                PDU_2.Text = aLAS8PDUS(17)
                PDU_3.Text = aLAS8PDUS(18)
                PDU_4.Text = aLAS8PDUS(19)
                PDU_5.Text = aLAS8PDUS(20)
                PDU_6.Text = aLAS8PDUS(21)
                PDU_7.Text = aLAS8PDUS(22)
                PDU_8.Text = aLAS8PDUS(23)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 65C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(16)
                PDU_2.Text = aLAS8PDUS(17)
                PDU_3.Text = aLAS8PDUS(18)
                PDU_4.Text = aLAS8PDUS(19)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 66C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(20)
                PDU_2.Text = aLAS8PDUS(21)
                PDU_3.Text = aLAS8PDUS(22)
                PDU_4.Text = aLAS8PDUS(23)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 71C" And cboEquipmentID_2.value = _
                "UPS 72C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(87)
                PDU_2.Text = aLAS8PDUS(88)
                PDU_3.Text = aLAS8PDUS(89)
                PDU_4.Text = aLAS8PDUS(90)
                PDU_5.Text = aLAS8PDUS(91)
                PDU_6.Text = aLAS8PDUS(92)
                PDU_7.Text = aLAS8PDUS(93)
                PDU_8.Text = aLAS8PDUS(94)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 71C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(87)
                PDU_2.Text = aLAS8PDUS(88)
                PDU_3.Text = aLAS8PDUS(89)
                PDU_4.Text = aLAS8PDUS(90)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 72C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(91)
                PDU_2.Text = aLAS8PDUS(92)
                PDU_3.Text = aLAS8PDUS(93)
                PDU_4.Text = aLAS8PDUS(94)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 77C" And cboEquipmentID_2.value = _
                "UPS 78C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(95)
                PDU_2.Text = aLAS8PDUS(96)
                PDU_3.Text = aLAS8PDUS(97)
                PDU_4.Text = aLAS8PDUS(98)
                PDU_5.Text = aLAS8PDUS(99)
                PDU_6.Text = aLAS8PDUS(100)
                PDU_7.Text = aLAS8PDUS(101)
                PDU_8.Text = aLAS8PDUS(102)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 77C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(95)
                PDU_2.Text = aLAS8PDUS(96)
                PDU_3.Text = aLAS8PDUS(97)
                PDU_4.Text = aLAS8PDUS(98)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 78C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(99)
                PDU_2.Text = aLAS8PDUS(100)
                PDU_3.Text = aLAS8PDUS(101)
                PDU_4.Text = aLAS8PDUS(102)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 83C" And cboEquipmentID_2.value = _
                "UPS 84C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(103)
                PDU_2.Text = aLAS8PDUS(104)
                PDU_3.Text = aLAS8PDUS(105)
                PDU_4.Text = aLAS8PDUS(106)
                PDU_5.Text = aLAS8PDUS(107)
                PDU_6.Text = aLAS8PDUS(108)
                PDU_7.Text = aLAS8PDUS(109)
                PDU_8.Text = aLAS8PDUS(110)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 83C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(103)
                PDU_2.Text = aLAS8PDUS(104)
                PDU_3.Text = aLAS8PDUS(105)
                PDU_4.Text = aLAS8PDUS(106)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 84C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS8PDUS(107)
                PDU_2.Text = aLAS8PDUS(108)
                PDU_3.Text = aLAS8PDUS(109)
                PDU_4.Text = aLAS8PDUS(110)
                
                 
            End If


 End Sub
  Private Sub setpduLAS9()
  'LAS 9 PDU List
    Dim PDU_1 As Range
    Dim PDU_2 As Range
    Dim PDU_3 As Range
    Dim PDU_4 As Range
    Dim PDU_5 As Range
    Dim PDU_6 As Range
    Dim PDU_7 As Range
    Dim PDU_8 As Range
    Dim aLAS9PDUS(0 To 89) As String

    
        aLAS9PDUS(0) = "PDU 361A"
        aLAS9PDUS(1) = "PDU 365A"
        aLAS9PDUS(2) = "PDU 368A"
        aLAS9PDUS(3) = "PDU 412A"
        aLAS9PDUS(4) = "PDU 364A"
        aLAS9PDUS(5) = "PDU 414A"
        aLAS9PDUS(6) = "PDU 415A"
        aLAS9PDUS(7) = "PDU 416A"
        aLAS9PDUS(8) = "PDU 369B"
        aLAS9PDUS(9) = "PDU 373B"
        aLAS9PDUS(10) = "PDU 376B"
        aLAS9PDUS(11) = "PDU 420B"
        aLAS9PDUS(12) = "PDU 372B"
        aLAS9PDUS(13) = "PDU 422B"
        aLAS9PDUS(14) = "PDU 423B"
        aLAS9PDUS(15) = "PDU 424B"
        aLAS9PDUS(16) = "PDU 377C"
        aLAS9PDUS(17) = "PDU 381C"
        aLAS9PDUS(18) = "PDU 384C"
        aLAS9PDUS(19) = "PDU 428C"
        aLAS9PDUS(20) = "PDU 380C"
        aLAS9PDUS(21) = "PDU 430C"
        aLAS9PDUS(22) = "PDU 431C"
        aLAS9PDUS(23) = "PDU 432C"
        aLAS9PDUS(24) = "PDU 386A"
        aLAS9PDUS(25) = "PDU 387A"
        aLAS9PDUS(26) = "PDU 388A"
        aLAS9PDUS(27) = "PDU 389A"
        aLAS9PDUS(28) = "PDU 390A"
        aLAS9PDUS(29) = "PDU 391A"
        aLAS9PDUS(30) = "PDU 392A"
        aLAS9PDUS(31) = "PDU 394B"
        aLAS9PDUS(32) = "PDU 395B"
        aLAS9PDUS(33) = "PDU 396B"
        aLAS9PDUS(34) = "PDU 397B"
        aLAS9PDUS(35) = "PDU 398B"
        aLAS9PDUS(36) = "PDU 399B"
        aLAS9PDUS(37) = "PDU 400B"
        aLAS9PDUS(38) = "PDU 402C"
        aLAS9PDUS(39) = "PDU 403C"
        aLAS9PDUS(40) = "PDU 404C"
        aLAS9PDUS(41) = "PDU 405C"
        aLAS9PDUS(42) = "PDU 406C"
        aLAS9PDUS(43) = "PDU 407C"
        aLAS9PDUS(44) = "PDU 408C"
        aLAS9PDUS(45) = "PDU 362A"
        aLAS9PDUS(46) = "PDU 363A"
        aLAS9PDUS(47) = "PDU 409A"
        aLAS9PDUS(48) = "PDU 410A"
        aLAS9PDUS(49) = "PDU 366A"
        aLAS9PDUS(50) = "PDU 367A"
        aLAS9PDUS(51) = "PDU 385A"
        aLAS9PDUS(52) = "PDU 370B"
        aLAS9PDUS(53) = "PDU 371B"
        aLAS9PDUS(54) = "PDU 417B"
        aLAS9PDUS(55) = "PDU 418B"
        aLAS9PDUS(56) = "PDU 374B"
        aLAS9PDUS(57) = "PDU 375B"
        aLAS9PDUS(58) = "PDU 393B"
        aLAS9PDUS(59) = "PDU 378C"
        aLAS9PDUS(60) = "PDU 379C"
        aLAS9PDUS(61) = "PDU 425C"
        aLAS9PDUS(62) = "PDU 426C"
        aLAS9PDUS(63) = "PDU 382C"
        aLAS9PDUS(64) = "PDU 383C"
        aLAS9PDUS(65) = "PDU 401C"
        aLAS9PDUS(66) = "PDU 337A"
        aLAS9PDUS(67) = "PDU 338A"
        aLAS9PDUS(68) = "PDU 339A"
        aLAS9PDUS(69) = "PDU 340A"
        aLAS9PDUS(70) = "PDU 341A"
        aLAS9PDUS(71) = "PDU 342A"
        aLAS9PDUS(72) = "PDU 343A"
        aLAS9PDUS(73) = "PDU 344A"
        aLAS9PDUS(74) = "PDU 345B"
        aLAS9PDUS(75) = "PDU 346B"
        aLAS9PDUS(76) = "PDU 347B"
        aLAS9PDUS(77) = "PDU 348B"
        aLAS9PDUS(78) = "PDU 349B"
        aLAS9PDUS(79) = "PDU 350B"
        aLAS9PDUS(80) = "PDU 351B"
        aLAS9PDUS(81) = "PDU 352B"
        aLAS9PDUS(82) = "PDU 353C"
        aLAS9PDUS(83) = "PDU 354C"
        aLAS9PDUS(84) = "PDU 355C"
        aLAS9PDUS(85) = "PDU 356C"
        aLAS9PDUS(86) = "PDU 357C"
        aLAS9PDUS(87) = "PDU 358C"
        aLAS9PDUS(88) = "PDU 359C"
        aLAS9PDUS(89) = "PDU 360C"

        
            Set PDU_1 = ActiveDocument.Bookmarks("tPDU_1").Range
            Set PDU_2 = ActiveDocument.Bookmarks("tPDU_2").Range
            Set PDU_3 = ActiveDocument.Bookmarks("tPDU_3").Range
            Set PDU_4 = ActiveDocument.Bookmarks("tPDU_4").Range
            Set PDU_5 = ActiveDocument.Bookmarks("tPDU_5").Range
            Set PDU_6 = ActiveDocument.Bookmarks("tPDU_6").Range
            Set PDU_7 = ActiveDocument.Bookmarks("tPDU_7").Range
            Set PDU_8 = ActiveDocument.Bookmarks("tPDU_8").Range

        'Setting the PDU's for any of the Power Systems Selected for NAP8
        

                If cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 103A" And cboEquipmentID_2.value = _
                "UPS 104A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(0)
                PDU_2.Text = aLAS9PDUS(1)
                PDU_3.Text = aLAS9PDUS(2)
                PDU_4.Text = aLAS9PDUS(3)
                PDU_5.Text = aLAS9PDUS(4)
                PDU_6.Text = aLAS9PDUS(5)
                PDU_7.Text = aLAS9PDUS(6)
                PDU_8.Text = aLAS9PDUS(7)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 103A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(0)
                PDU_2.Text = aLAS9PDUS(1)
                PDU_3.Text = aLAS9PDUS(2)
                PDU_4.Text = aLAS9PDUS(3)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 104A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(4)
                PDU_2.Text = aLAS9PDUS(5)
                PDU_3.Text = aLAS9PDUS(6)
                PDU_4.Text = aLAS9PDUS(7)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 97A" And cboEquipmentID_2.value = _
                "UPS 98A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(24)
                PDU_2.Text = aLAS9PDUS(25)
                PDU_3.Text = aLAS9PDUS(26)
                PDU_4.Text = aLAS9PDUS(27)
                PDU_5.Text = aLAS9PDUS(28)
                PDU_6.Text = aLAS9PDUS(29)
                PDU_7.Text = aLAS9PDUS(30)
            
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 97A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(24)
                PDU_2.Text = aLAS9PDUS(25)
                PDU_3.Text = aLAS9PDUS(26)

                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 98A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(27)
                PDU_2.Text = aLAS9PDUS(28)
                PDU_3.Text = aLAS9PDUS(29)
                PDU_4.Text = aLAS9PDUS(30)

                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 91A" And cboEquipmentID_2.value = _
                "UPS 92A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(45)
                PDU_2.Text = aLAS9PDUS(46)
                PDU_3.Text = aLAS9PDUS(47)
                PDU_4.Text = aLAS9PDUS(48)
                PDU_5.Text = aLAS9PDUS(49)
                PDU_6.Text = aLAS9PDUS(50)
                PDU_7.Text = aLAS9PDUS(51)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 91A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(45)
                PDU_2.Text = aLAS9PDUS(46)
                PDU_3.Text = aLAS9PDUS(47)
                PDU_4.Text = aLAS9PDUS(48)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 92A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(49)
                PDU_2.Text = aLAS9PDUS(50)
                PDU_3.Text = aLAS9PDUS(51)

                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 85A" And cboEquipmentID_2.value = _
                "UPS 86A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(66)
                PDU_2.Text = aLAS9PDUS(67)
                PDU_3.Text = aLAS9PDUS(68)
                PDU_4.Text = aLAS9PDUS(69)
                PDU_5.Text = aLAS9PDUS(70)
                PDU_6.Text = aLAS9PDUS(71)
                PDU_7.Text = aLAS9PDUS(72)
                PDU_8.Text = aLAS9PDUS(73)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 85A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(66)
                PDU_2.Text = aLAS9PDUS(67)
                PDU_3.Text = aLAS9PDUS(68)
                PDU_4.Text = aLAS9PDUS(69)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 86A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(70)
                PDU_2.Text = aLAS9PDUS(71)
                PDU_3.Text = aLAS9PDUS(72)
                PDU_4.Text = aLAS9PDUS(73)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 105B" And cboEquipmentID_2.value = _
                "UPS 106B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(8)
                PDU_2.Text = aLAS9PDUS(9)
                PDU_3.Text = aLAS9PDUS(10)
                PDU_4.Text = aLAS9PDUS(11)
                PDU_5.Text = aLAS9PDUS(12)
                PDU_6.Text = aLAS9PDUS(13)
                PDU_7.Text = aLAS9PDUS(14)
                PDU_8.Text = aLAS9PDUS(15)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 105B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(8)
                PDU_2.Text = aLAS9PDUS(9)
                PDU_3.Text = aLAS9PDUS(10)
                PDU_4.Text = aLAS9PDUS(11)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 106B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(12)
                PDU_2.Text = aLAS9PDUS(13)
                PDU_3.Text = aLAS9PDUS(14)
                PDU_4.Text = aLAS9PDUS(15)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 99B" And cboEquipmentID_2.value = _
                "UPS 100B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(31)
                PDU_2.Text = aLAS9PDUS(32)
                PDU_3.Text = aLAS9PDUS(33)
                PDU_4.Text = aLAS9PDUS(34)
                PDU_5.Text = aLAS9PDUS(35)
                PDU_6.Text = aLAS9PDUS(36)
                PDU_7.Text = aLAS9PDUS(37)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 99B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(31)
                PDU_2.Text = aLAS9PDUS(32)
                PDU_3.Text = aLAS9PDUS(33)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 100B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(34)
                PDU_2.Text = aLAS9PDUS(35)
                PDU_3.Text = aLAS9PDUS(36)
                PDU_4.Text = aLAS9PDUS(37)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 93B" And cboEquipmentID_2.value = _
                "UPS 94B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(52)
                PDU_2.Text = aLAS9PDUS(53)
                PDU_3.Text = aLAS9PDUS(54)
                PDU_4.Text = aLAS9PDUS(55)
                PDU_5.Text = aLAS9PDUS(56)
                PDU_6.Text = aLAS9PDUS(57)
                PDU_7.Text = aLAS9PDUS(58)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 93B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(52)
                PDU_2.Text = aLAS9PDUS(53)
                PDU_3.Text = aLAS9PDUS(54)
                PDU_4.Text = aLAS9PDUS(55)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 94B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(56)
                PDU_2.Text = aLAS9PDUS(57)
                PDU_3.Text = aLAS9PDUS(58)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 87B" And cboEquipmentID_2.value = _
                "UPS 88B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(74)
                PDU_2.Text = aLAS9PDUS(75)
                PDU_3.Text = aLAS9PDUS(76)
                PDU_4.Text = aLAS9PDUS(77)
                PDU_5.Text = aLAS9PDUS(78)
                PDU_6.Text = aLAS9PDUS(79)
                PDU_7.Text = aLAS9PDUS(80)
                PDU_8.Text = aLAS9PDUS(81)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 87B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(74)
                PDU_2.Text = aLAS9PDUS(75)
                PDU_3.Text = aLAS9PDUS(76)
                PDU_4.Text = aLAS9PDUS(77)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 88B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(78)
                PDU_2.Text = aLAS9PDUS(79)
                PDU_3.Text = aLAS9PDUS(80)
                PDU_4.Text = aLAS9PDUS(81)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 107C" And cboEquipmentID_2.value = _
                "UPS 108C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(16)
                PDU_2.Text = aLAS9PDUS(17)
                PDU_3.Text = aLAS9PDUS(18)
                PDU_4.Text = aLAS9PDUS(19)
                PDU_5.Text = aLAS9PDUS(20)
                PDU_6.Text = aLAS9PDUS(21)
                PDU_7.Text = aLAS9PDUS(22)
                PDU_8.Text = aLAS9PDUS(23)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 107C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(16)
                PDU_2.Text = aLAS9PDUS(17)
                PDU_3.Text = aLAS9PDUS(18)
                PDU_4.Text = aLAS9PDUS(19)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 108C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(20)
                PDU_2.Text = aLAS9PDUS(21)
                PDU_3.Text = aLAS9PDUS(22)
                PDU_4.Text = aLAS9PDUS(23)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 101C" And cboEquipmentID_2.value = _
                "UPS 102C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(38)
                PDU_2.Text = aLAS9PDUS(39)
                PDU_3.Text = aLAS9PDUS(40)
                PDU_4.Text = aLAS9PDUS(41)
                PDU_5.Text = aLAS9PDUS(42)
                PDU_6.Text = aLAS9PDUS(43)
                PDU_7.Text = aLAS9PDUS(44)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 101C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(38)
                PDU_2.Text = aLAS9PDUS(39)
                PDU_3.Text = aLAS9PDUS(40)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 102C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(41)
                PDU_2.Text = aLAS9PDUS(42)
                PDU_3.Text = aLAS9PDUS(43)
                PDU_4.Text = aLAS9PDUS(44)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 95C" And cboEquipmentID_2.value = _
                "UPS 96C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(59)
                PDU_2.Text = aLAS9PDUS(60)
                PDU_3.Text = aLAS9PDUS(61)
                PDU_4.Text = aLAS9PDUS(62)
                PDU_5.Text = aLAS9PDUS(63)
                PDU_6.Text = aLAS9PDUS(64)
                PDU_7.Text = aLAS9PDUS(65)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 95C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(59)
                PDU_2.Text = aLAS9PDUS(60)
                PDU_3.Text = aLAS9PDUS(61)
                PDU_4.Text = aLAS9PDUS(62)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 96C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(63)
                PDU_2.Text = aLAS9PDUS(64)
                PDU_3.Text = aLAS9PDUS(65)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 89C" And cboEquipmentID_2.value = _
                "UPS 90C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(82)
                PDU_2.Text = aLAS9PDUS(83)
                PDU_3.Text = aLAS9PDUS(84)
                PDU_4.Text = aLAS9PDUS(85)
                PDU_5.Text = aLAS9PDUS(86)
                PDU_6.Text = aLAS9PDUS(87)
                PDU_7.Text = aLAS9PDUS(88)
                PDU_8.Text = aLAS9PDUS(89)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 89C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(82)
                PDU_2.Text = aLAS9PDUS(83)
                PDU_3.Text = aLAS9PDUS(84)
                PDU_4.Text = aLAS9PDUS(85)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 90C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS9PDUS(86)
                PDU_2.Text = aLAS9PDUS(87)
                PDU_3.Text = aLAS9PDUS(88)
                PDU_4.Text = aLAS9PDUS(89)
                
                 
            End If


 End Sub
   Private Sub setpduLAS10()
  'LAS 10 PDU List
    Dim PDU_1 As Range
    Dim PDU_2 As Range
    Dim PDU_3 As Range
    Dim PDU_4 As Range
    Dim PDU_5 As Range
    Dim PDU_6 As Range
    Dim PDU_7 As Range
    Dim PDU_8 As Range
    Dim aLAS10PDUS(0 To 94) As String

    
        aLAS10PDUS(0) = "PDU 433A"
        aLAS10PDUS(1) = "PDU 434A"
        aLAS10PDUS(2) = "PDU 435A"
        aLAS10PDUS(3) = "PDU 457A"
        aLAS10PDUS(4) = "PDU 437A"
        aLAS10PDUS(5) = "PDU 438A"
        aLAS10PDUS(6) = "PDU 461A"
        aLAS10PDUS(7) = "PDU "
        aLAS10PDUS(8) = "PDU 441B"
        aLAS10PDUS(9) = "PDU 442B"
        aLAS10PDUS(10) = "PDU 443B"
        aLAS10PDUS(11) = "PDU 465B"
        aLAS10PDUS(12) = "PDU 445B"
        aLAS10PDUS(13) = "PDU 446B"
        aLAS10PDUS(14) = "PDU 469B"
        aLAS10PDUS(15) = "PDU "
        aLAS10PDUS(16) = "PDU 449C"
        aLAS10PDUS(17) = "PDU 450C"
        aLAS10PDUS(18) = "PDU "
        aLAS10PDUS(19) = "PDU "
        aLAS10PDUS(20) = "PDU 453C"
        aLAS10PDUS(21) = "PDU 454C"
        aLAS10PDUS(22) = "PDU 477C"
        aLAS10PDUS(23) = "PDU "
        aLAS10PDUS(24) = "PDU "
        aLAS10PDUS(25) = "PDU "
        aLAS10PDUS(26) = "PDU "
        aLAS10PDUS(27) = "PDU "
        aLAS10PDUS(28) = "PDU "
        aLAS10PDUS(29) = "PDU "
        aLAS10PDUS(30) = "PDU "
        aLAS10PDUS(31) = "PDU "
        aLAS10PDUS(32) = "PDU "
        aLAS10PDUS(33) = "PDU "
        aLAS10PDUS(34) = "PDU "
        aLAS10PDUS(35) = "PDU "
        aLAS10PDUS(36) = "PDU "
        aLAS10PDUS(37) = "PDU "
        aLAS10PDUS(38) = "PDU "
        aLAS10PDUS(39) = "PDU "
        aLAS10PDUS(40) = "PDU "
        aLAS10PDUS(41) = "PDU "
        aLAS10PDUS(42) = "PDU "
        aLAS10PDUS(43) = "PDU "
        aLAS10PDUS(44) = "PDU "
        aLAS10PDUS(45) = "PDU "
        aLAS10PDUS(46) = "PDU "
        aLAS10PDUS(47) = "PDU "
        aLAS10PDUS(48) = "PDU "
        aLAS10PDUS(49) = "PDU "
        aLAS10PDUS(50) = "PDU "
        aLAS10PDUS(51) = "PDU "
        aLAS10PDUS(52) = "PDU "
        aLAS10PDUS(53) = "PDU "
        aLAS10PDUS(54) = "PDU "
        aLAS10PDUS(55) = "PDU "
        aLAS10PDUS(56) = "PDU "
        aLAS10PDUS(57) = "PDU "
        aLAS10PDUS(58) = "PDU "
        aLAS10PDUS(59) = "PDU "
        aLAS10PDUS(60) = "PDU "
        aLAS10PDUS(61) = "PDU "
        aLAS10PDUS(62) = "PDU "
        aLAS10PDUS(63) = "PDU "
        aLAS10PDUS(64) = "PDU "
        aLAS10PDUS(65) = "PDU "
        aLAS10PDUS(66) = "PDU "
        aLAS10PDUS(67) = "PDU "
        aLAS10PDUS(68) = "PDU "
        aLAS10PDUS(69) = "PDU "
        aLAS10PDUS(70) = "PDU "
        aLAS10PDUS(71) = "PDU "
        aLAS10PDUS(72) = "PDU "
        aLAS10PDUS(73) = "PDU "
        aLAS10PDUS(74) = "PDU "
        aLAS10PDUS(75) = "PDU "
        aLAS10PDUS(76) = "PDU "
        aLAS10PDUS(77) = "PDU "
        aLAS10PDUS(78) = "PDU "
        aLAS10PDUS(79) = "PDU "
        aLAS10PDUS(80) = "PDU "
        aLAS10PDUS(81) = "PDU "
        aLAS10PDUS(82) = "PDU "
        aLAS10PDUS(83) = "PDU "
        aLAS10PDUS(84) = "PDU "
        aLAS10PDUS(85) = "PDU "
        aLAS10PDUS(86) = "PDU "
        aLAS10PDUS(87) = "PDU "
        aLAS10PDUS(88) = "PDU "
        aLAS10PDUS(89) = "PDU "
        aLAS10PDUS(90) = "PDU "
        aLAS10PDUS(91) = "PDU "
        aLAS10PDUS(92) = "PDU "
        aLAS10PDUS(93) = "PDU "
        aLAS10PDUS(94) = "PDU "

        
            Set PDU_1 = ActiveDocument.Bookmarks("tPDU_1").Range
            Set PDU_2 = ActiveDocument.Bookmarks("tPDU_2").Range
            Set PDU_3 = ActiveDocument.Bookmarks("tPDU_3").Range
            Set PDU_4 = ActiveDocument.Bookmarks("tPDU_4").Range
            Set PDU_5 = ActiveDocument.Bookmarks("tPDU_5").Range
            Set PDU_6 = ActiveDocument.Bookmarks("tPDU_6").Range
            Set PDU_7 = ActiveDocument.Bookmarks("tPDU_7").Range
            Set PDU_8 = ActiveDocument.Bookmarks("tPDU_8").Range

        'Setting the PDU's for any of the Power Systems Selected for NAP 10
        

                If cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 1A" And cboEquipmentID_2.value = _
                "UPS 2A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(0)
                PDU_2.Text = aLAS10PDUS(1)
                PDU_3.Text = aLAS10PDUS(2)
                PDU_4.Text = aLAS10PDUS(3)
                PDU_5.Text = aLAS10PDUS(4)
                PDU_6.Text = aLAS10PDUS(5)
                PDU_7.Text = aLAS10PDUS(6)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 1A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(0)
                PDU_2.Text = aLAS10PDUS(1)
                PDU_3.Text = aLAS10PDUS(2)
                PDU_4.Text = aLAS10PDUS(3)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 2A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_5.Text = aLAS10PDUS(4)
                PDU_6.Text = aLAS10PDUS(5)
                PDU_7.Text = aLAS10PDUS(6)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 3B" And cboEquipmentID_2.value = _
                "UPS 4B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(8)
                PDU_2.Text = aLAS10PDUS(9)
                PDU_3.Text = aLAS10PDUS(10)
                PDU_4.Text = aLAS10PDUS(11)
                PDU_5.Text = aLAS10PDUS(12)
                PDU_6.Text = aLAS10PDUS(13)
                PDU_7.Text = aLAS10PDUS(14)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 3B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(8)
                PDU_2.Text = aLAS10PDUS(9)
                PDU_3.Text = aLAS10PDUS(10)
                PDU_4.Text = aLAS10PDUS(11)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 4B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(12)
                PDU_2.Text = aLAS10PDUS(13)
                PDU_3.Text = aLAS10PDUS(14)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 5C" And cboEquipmentID_2.value = _
                "UPS 6C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(16)
                PDU_2.Text = aLAS10PDUS(17)
                PDU_5.Text = aLAS10PDUS(20)
                PDU_6.Text = aLAS10PDUS(21)
                PDU_7.Text = aLAS10PDUS(22)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 5C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(16)
                PDU_2.Text = aLAS10PDUS(17)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 6C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS10PDUS(20)
                PDU_2.Text = aLAS10PDUS(21)
                PDU_3.Text = aLAS10PDUS(22)
                 
            End If


 End Sub
    Private Sub setpduLAS11()
  'LAS 11 PDU List
    Dim PDU_1 As Range
    Dim PDU_2 As Range
    Dim PDU_3 As Range
    Dim PDU_4 As Range
    Dim PDU_5 As Range
    Dim PDU_6 As Range
    Dim PDU_7 As Range
    Dim PDU_8 As Range
    Dim aLAS11PDUS(0 To 110) As String

    
        aLAS11PDUS(0) = "PDU 241A"
        aLAS11PDUS(1) = "PDU 242A"
        aLAS11PDUS(2) = "PDU 243A"
        aLAS11PDUS(3) = "PDU 244A"
        aLAS11PDUS(4) = "PDU 245A"
        aLAS11PDUS(5) = "PDU 246A"
        aLAS11PDUS(6) = "PDU 247A"
        aLAS11PDUS(7) = "PDU 248A"
        aLAS11PDUS(8) = "PDU 249B"
        aLAS11PDUS(9) = "PDU 250B"
        aLAS11PDUS(10) = "PDU 251B"
        aLAS11PDUS(11) = "PDU 252B"
        aLAS11PDUS(12) = "PDU 253B"
        aLAS11PDUS(13) = "PDU 254B"
        aLAS11PDUS(14) = "PDU 255B"
        aLAS11PDUS(15) = "PDU 256B"
        aLAS11PDUS(16) = "PDU 257C"
        aLAS11PDUS(17) = "PDU 258C"
        aLAS11PDUS(18) = "PDU 259C"
        aLAS11PDUS(19) = "PDU 260C"
        aLAS11PDUS(20) = "PDU 261C"
        aLAS11PDUS(21) = "PDU 262C"
        aLAS11PDUS(22) = "PDU 264C"
        aLAS11PDUS(23) = "PDU 265A"
        aLAS11PDUS(24) = "PDU 266A"
        aLAS11PDUS(25) = "PDU 267A"
        aLAS11PDUS(26) = "PDU 268A"
        aLAS11PDUS(27) = "PDU 269A"
        aLAS11PDUS(28) = "PDU 270A"
        aLAS11PDUS(29) = "PDU 271A"
        aLAS11PDUS(30) = "PDU 272A"
        aLAS11PDUS(31) = "PDU 289A"
        aLAS11PDUS(32) = "PDU 290A"
        aLAS11PDUS(33) = "PDU 291A"
        aLAS11PDUS(34) = "PDU 292A"
        aLAS11PDUS(35) = "PDU 293A"
        aLAS11PDUS(36) = "PDU 294A"
        aLAS11PDUS(37) = "PDU 295A"
        aLAS11PDUS(38) = "PDU 296A"
        aLAS11PDUS(39) = "PDU 313A"
        aLAS11PDUS(40) = "PDU 314A"
        aLAS11PDUS(41) = "PDU 315A"
        aLAS11PDUS(42) = "PDU 316A"
        aLAS11PDUS(43) = "PDU 317A"
        aLAS11PDUS(44) = "PDU 318A"
        aLAS11PDUS(45) = "PDU 319A"
        aLAS11PDUS(46) = "PDU 320A"
        aLAS11PDUS(47) = "PDU 249B"
        aLAS11PDUS(48) = "PDU 250B"
        aLAS11PDUS(49) = "PDU 251B"
        aLAS11PDUS(50) = "PDU 252B"
        aLAS11PDUS(51) = "PDU 253B"
        aLAS11PDUS(52) = "PDU 254B"
        aLAS11PDUS(53) = "PDU 255B"
        aLAS11PDUS(54) = "PDU 256B"
        aLAS11PDUS(55) = "PDU 273B"
        aLAS11PDUS(56) = "PDU 274B"
        aLAS11PDUS(57) = "PDU 275B"
        aLAS11PDUS(58) = "PDU 276B"
        aLAS11PDUS(59) = "PDU 277B"
        aLAS11PDUS(60) = "PDU 278B"
        aLAS11PDUS(61) = "PDU 279B"
        aLAS11PDUS(62) = "PDU 280B"
        aLAS11PDUS(63) = "PDU 297B"
        aLAS11PDUS(64) = "PDU 298B"
        aLAS11PDUS(65) = "PDU 299B"
        aLAS11PDUS(66) = "PDU 300B"
        aLAS11PDUS(67) = "PDU 301B"
        aLAS11PDUS(68) = "PDU 302B"
        aLAS11PDUS(69) = "PDU 303B"
        aLAS11PDUS(70) = "PDU 304B"
        aLAS11PDUS(71) = "PDU 321B"
        aLAS11PDUS(72) = "PDU 322B"
        aLAS11PDUS(73) = "PDU 323B"
        aLAS11PDUS(74) = "PDU 324B"
        aLAS11PDUS(75) = "PDU 325B"
        aLAS11PDUS(76) = "PDU 326B"
        aLAS11PDUS(77) = "PDU 327B"
        aLAS11PDUS(78) = "PDU 328B"
        aLAS11PDUS(79) = "PDU 257C"
        aLAS11PDUS(80) = "PDU 258C"
        aLAS11PDUS(81) = "PDU 259C"
        aLAS11PDUS(82) = "PDU 260C"
        aLAS11PDUS(83) = "PDU 261C"
        aLAS11PDUS(84) = "PDU 262C"
        aLAS11PDUS(85) = "PDU 263C"
        aLAS11PDUS(86) = "PDU 264C"
        aLAS11PDUS(87) = "PDU 281C"
        aLAS11PDUS(88) = "PDU 282C"
        aLAS11PDUS(89) = "PDU 283C"
        aLAS11PDUS(90) = "PDU 284C"
        aLAS11PDUS(91) = "PDU 285C"
        aLAS11PDUS(92) = "PDU 286C"
        aLAS11PDUS(93) = "PDU 287C"
        aLAS11PDUS(94) = "PDU 288C"
        aLAS11PDUS(95) = "PDU 305C"
        aLAS11PDUS(96) = "PDU 306C"
        aLAS11PDUS(97) = "PDU 307C"
        aLAS11PDUS(98) = "PDU 308C"
        aLAS11PDUS(99) = "PDU 309C"
        aLAS11PDUS(100) = "PDU 310C"
        aLAS11PDUS(101) = "PDU 311C"
        aLAS11PDUS(102) = "PDU 312C"
        aLAS11PDUS(103) = "PDU 329C"
        aLAS11PDUS(104) = "PDU 330C"
        aLAS11PDUS(105) = "PDU 331C"
        aLAS11PDUS(106) = "PDU 332C"
        aLAS11PDUS(107) = "PDU 333C"
        aLAS11PDUS(108) = "PDU 334C"
        aLAS11PDUS(109) = "PDU 335C"
        aLAS11PDUS(110) = "PDU 336C"
        
            Set PDU_1 = ActiveDocument.Bookmarks("tPDU_1").Range
            Set PDU_2 = ActiveDocument.Bookmarks("tPDU_2").Range
            Set PDU_3 = ActiveDocument.Bookmarks("tPDU_3").Range
            Set PDU_4 = ActiveDocument.Bookmarks("tPDU_4").Range
            Set PDU_5 = ActiveDocument.Bookmarks("tPDU_5").Range
            Set PDU_6 = ActiveDocument.Bookmarks("tPDU_6").Range
            Set PDU_7 = ActiveDocument.Bookmarks("tPDU_7").Range
            Set PDU_8 = ActiveDocument.Bookmarks("tPDU_8").Range

        'Setting the PDU's for any of the Power Systems Selected for NAP8
        

                If cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 61A" And cboEquipmentID_2.value = _
                "UPS 62A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(0)
                PDU_2.Text = aLAS11PDUS(1)
                PDU_3.Text = aLAS11PDUS(2)
                PDU_4.Text = aLAS11PDUS(3)
                PDU_5.Text = aLAS11PDUS(4)
                PDU_6.Text = aLAS11PDUS(5)
                PDU_7.Text = aLAS11PDUS(6)
                PDU_8.Text = aLAS11PDUS(7)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 61A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(0)
                PDU_2.Text = aLAS11PDUS(1)
                PDU_3.Text = aLAS11PDUS(2)
                PDU_4.Text = aLAS11PDUS(3)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 62A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(4)
                PDU_2.Text = aLAS11PDUS(5)
                PDU_3.Text = aLAS11PDUS(6)
                PDU_4.Text = aLAS11PDUS(7)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 67A" And cboEquipmentID_2.value = _
                "UPS 68A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(23)
                PDU_2.Text = aLAS11PDUS(24)
                PDU_3.Text = aLAS11PDUS(25)
                PDU_4.Text = aLAS11PDUS(26)
                PDU_5.Text = aLAS11PDUS(27)
                PDU_6.Text = aLAS11PDUS(28)
                PDU_7.Text = aLAS11PDUS(29)
                PDU_8.Text = aLAS11PDUS(30)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 67A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(23)
                PDU_2.Text = aLAS11PDUS(24)
                PDU_3.Text = aLAS11PDUS(25)
                PDU_4.Text = aLAS11PDUS(26)

                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 68A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(27)
                PDU_2.Text = aLAS11PDUS(28)
                PDU_3.Text = aLAS11PDUS(29)
                PDU_4.Text = aLAS11PDUS(30)

                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 73A" And cboEquipmentID_2.value = _
                "UPS 74A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(31)
                PDU_2.Text = aLAS11PDUS(32)
                PDU_3.Text = aLAS11PDUS(33)
                PDU_4.Text = aLAS11PDUS(34)
                PDU_5.Text = aLAS11PDUS(35)
                PDU_6.Text = aLAS11PDUS(36)
                PDU_7.Text = aLAS11PDUS(37)
                PDU_8.Text = aLAS11PDUS(38)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 73A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(31)
                PDU_2.Text = aLAS11PDUS(32)
                PDU_3.Text = aLAS11PDUS(33)
                PDU_4.Text = aLAS11PDUS(34)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 74A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(35)
                PDU_2.Text = aLAS11PDUS(36)
                PDU_3.Text = aLAS11PDUS(37)
                PDU_4.Text = aLAS11PDUS(38)

                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 79A" And cboEquipmentID_2.value = _
                "UPS 80A" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(39)
                PDU_2.Text = aLAS11PDUS(40)
                PDU_3.Text = aLAS11PDUS(41)
                PDU_4.Text = aLAS11PDUS(42)
                PDU_5.Text = aLAS11PDUS(43)
                PDU_6.Text = aLAS11PDUS(44)
                PDU_7.Text = aLAS11PDUS(45)
                PDU_8.Text = aLAS11PDUS(46)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 79A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(39)
                PDU_2.Text = aLAS11PDUS(40)
                PDU_3.Text = aLAS11PDUS(41)
                PDU_4.Text = aLAS11PDUS(42)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 80A" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(43)
                PDU_2.Text = aLAS11PDUS(44)
                PDU_3.Text = aLAS11PDUS(45)
                PDU_4.Text = aLAS11PDUS(46)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 63B" And cboEquipmentID_2.value = _
                "UPS 64B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(8)
                PDU_2.Text = aLAS11PDUS(9)
                PDU_3.Text = aLAS11PDUS(10)
                PDU_4.Text = aLAS11PDUS(11)
                PDU_5.Text = aLAS11PDUS(12)
                PDU_6.Text = aLAS11PDUS(13)
                PDU_7.Text = aLAS11PDUS(14)
                PDU_8.Text = aLAS11PDUS(15)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 63B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(8)
                PDU_2.Text = aLAS11PDUS(9)
                PDU_3.Text = aLAS11PDUS(10)
                PDU_4.Text = aLAS11PDUS(11)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 64B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(12)
                PDU_2.Text = aLAS11PDUS(13)
                PDU_3.Text = aLAS11PDUS(14)
                PDU_4.Text = aLAS11PDUS(15)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 69B" And cboEquipmentID_2.value = _
                "UPS 70B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(55)
                PDU_2.Text = aLAS11PDUS(56)
                PDU_3.Text = aLAS11PDUS(57)
                PDU_4.Text = aLAS11PDUS(58)
                PDU_5.Text = aLAS11PDUS(59)
                PDU_6.Text = aLAS11PDUS(60)
                PDU_7.Text = aLAS11PDUS(61)
                PDU_8.Text = aLAS11PDUS(62)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 69B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(55)
                PDU_2.Text = aLAS11PDUS(56)
                PDU_3.Text = aLAS11PDUS(57)
                PDU_4.Text = aLAS11PDUS(58)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 70B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(59)
                PDU_2.Text = aLAS11PDUS(60)
                PDU_3.Text = aLAS11PDUS(61)
                PDU_4.Text = aLAS11PDUS(62)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 75B" And cboEquipmentID_2.value = _
                "UPS 76B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(63)
                PDU_2.Text = aLAS11PDUS(64)
                PDU_3.Text = aLAS11PDUS(65)
                PDU_4.Text = aLAS11PDUS(66)
                PDU_5.Text = aLAS11PDUS(67)
                PDU_6.Text = aLAS11PDUS(68)
                PDU_7.Text = aLAS11PDUS(69)
                PDU_8.Text = aLAS11PDUS(70)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 75B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(63)
                PDU_2.Text = aLAS11PDUS(64)
                PDU_3.Text = aLAS11PDUS(65)
                PDU_4.Text = aLAS11PDUS(66)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 76B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(67)
                PDU_2.Text = aLAS11PDUS(68)
                PDU_3.Text = aLAS11PDUS(69)
                PDU_4.Text = aLAS11PDUS(70)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 81B" And cboEquipmentID_2.value = _
                "UPS 82B" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(71)
                PDU_2.Text = aLAS11PDUS(72)
                PDU_3.Text = aLAS11PDUS(73)
                PDU_4.Text = aLAS11PDUS(74)
                PDU_5.Text = aLAS11PDUS(75)
                PDU_6.Text = aLAS11PDUS(76)
                PDU_7.Text = aLAS11PDUS(77)
                PDU_8.Text = aLAS11PDUS(78)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 81B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(71)
                PDU_2.Text = aLAS11PDUS(72)
                PDU_3.Text = aLAS11PDUS(73)
                PDU_4.Text = aLAS11PDUS(74)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 82B" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(75)
                PDU_2.Text = aLAS11PDUS(76)
                PDU_3.Text = aLAS11PDUS(77)
                PDU_4.Text = aLAS11PDUS(78)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 65C" And cboEquipmentID_2.value = _
                "UPS 66C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(16)
                PDU_2.Text = aLAS11PDUS(17)
                PDU_3.Text = aLAS11PDUS(18)
                PDU_4.Text = aLAS11PDUS(19)
                PDU_5.Text = aLAS11PDUS(20)
                PDU_6.Text = aLAS11PDUS(21)
                PDU_7.Text = aLAS11PDUS(22)
                PDU_8.Text = aLAS11PDUS(23)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 65C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(16)
                PDU_2.Text = aLAS11PDUS(17)
                PDU_3.Text = aLAS11PDUS(18)
                PDU_4.Text = aLAS11PDUS(19)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 66C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(20)
                PDU_2.Text = aLAS11PDUS(21)
                PDU_3.Text = aLAS11PDUS(22)
                PDU_4.Text = aLAS11PDUS(23)
            
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 71C" And cboEquipmentID_2.value = _
                "UPS 72C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(87)
                PDU_2.Text = aLAS11PDUS(88)
                PDU_3.Text = aLAS11PDUS(89)
                PDU_4.Text = aLAS11PDUS(90)
                PDU_5.Text = aLAS11PDUS(91)
                PDU_6.Text = aLAS11PDUS(92)
                PDU_7.Text = aLAS11PDUS(93)
                PDU_8.Text = aLAS11PDUS(94)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 71C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(87)
                PDU_2.Text = aLAS11PDUS(88)
                PDU_3.Text = aLAS11PDUS(89)
                PDU_4.Text = aLAS11PDUS(90)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 72C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(91)
                PDU_2.Text = aLAS11PDUS(92)
                PDU_3.Text = aLAS11PDUS(93)
                PDU_4.Text = aLAS11PDUS(94)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 77C" And cboEquipmentID_2.value = _
                "UPS 78C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(95)
                PDU_2.Text = aLAS11PDUS(96)
                PDU_3.Text = aLAS11PDUS(97)
                PDU_4.Text = aLAS11PDUS(98)
                PDU_5.Text = aLAS11PDUS(99)
                PDU_6.Text = aLAS11PDUS(100)
                PDU_7.Text = aLAS11PDUS(101)
                PDU_8.Text = aLAS11PDUS(102)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 77C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(95)
               PDU_2.Text = aLAS11PDUS(96)
                PDU_3.Text = aLAS11PDUS(97)
                PDU_4.Text = aLAS11PDUS(98)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 78C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(99)
                PDU_2.Text = aLAS11PDUS(100)
                PDU_3.Text = aLAS11PDUS(101)
                PDU_4.Text = aLAS11PDUS(102)
                
                ElseIf cbonumberofups.value = "2" And cboEquipmentID.value = "UPS 83C" And cboEquipmentID_2.value = _
                "UPS 84C" And (cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "2 UPS's Annual PM w/ Cal" Or cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(103)
                PDU_2.Text = aLAS11PDUS(104)
                PDU_3.Text = aLAS11PDUS(105)
                PDU_4.Text = aLAS11PDUS(106)
                PDU_5.Text = aLAS11PDUS(107)
                PDU_6.Text = aLAS11PDUS(108)
                PDU_7.Text = aLAS11PDUS(109)
                PDU_8.Text = aLAS11PDUS(110)
                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 83C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(103)
                PDU_2.Text = aLAS11PDUS(104)
                PDU_3.Text = aLAS11PDUS(105)
                PDU_4.Text = aLAS11PDUS(106)
                                
                ElseIf cbonumberofups.value = "1" And cboEquipmentID.value = "UPS 84C" And cboEquipmentID_2.value = _
                "" And (cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Or cbotypeofmaintenance.value = _
                "1 UPS Annual PM w/ Cal" Or cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Or _
                cbotypeofmaintenance.value = "1 UPS Corrective Maintenance") Then
                PDU_1.Text = aLAS11PDUS(107)
                PDU_2.Text = aLAS11PDUS(108)
                PDU_3.Text = aLAS11PDUS(109)
                PDU_4.Text = aLAS11PDUS(110)
                
                 
            End If


 End Sub

Public Sub OK_Click()
    
    Me.Hide
    Application.ScreenUpdating = False
    
    Dim Criticalitylevel As Range
    Dim Title As Range
    Dim BS1 As Range
    Dim BS1_1 As Range
    Dim GSBA As Range
    Dim GSBA_1 As Range
    Dim GSBA_2 As Range
    Dim GSBA_3 As Range
    Dim GSBA_4 As Range
    Dim GSBA_5 As Range
    Dim GSBB As Range
    Dim GSBB_1 As Range
    Dim GSBB_2 As Range
    Dim GSBB_3 As Range
    Dim GSBB_4 As Range
    Dim GSBB_5 As Range
    Dim GSBB_6 As Range
    Dim GSBC As Range
    Dim GSBC_1 As Range
    Dim GSBC_2 As Range
    Dim GSBC_3 As Range
    Dim GSBC_4 As Range
    Dim GSBC_5 As Range
    Dim MBS_GSB As Range
    Dim MVS As Range
    Dim MBS As Range
    Dim EquipmentID As Range
    Dim SPA As Range
    Dim spa_3 As Range
    Dim SPA_2 As Range
    Dim SPB As Range
    Dim SPB_2 As Range
    Dim SPB_3 As Range
    Dim spc_1 As Range
    Dim SPC_3 As Range
    Dim SPC_4 As Range
    Dim ISX As Range
    Dim ISX_2 As Range
    Dim UPS As Range
    Dim UPS_1 As Range
    Dim UPS_2 As Range
    Dim UPS_3 As Range
    Dim UPS_4 As Range
    Dim UPS_5 As Range
    Dim UPS_6 As Range
    Dim UPS_7 As Range
    Dim UPS_8 As Range
    Dim UPS_9 As Range
    Dim UPS_10 As Range
    Dim UPS_11 As Range
    Dim UPS_12 As Range
    Dim Site As Range
    Dim buildingName As Range
    Dim address As Range
    Dim upscolor As Range
    Dim i As Integer
    Dim referencedoc As Document
    Dim startpoint As Range
    Dim targetdoc As Document
    Dim targettable As table
    Dim Starttime As Range
    Dim Endtime As Range
    Dim Startdate As Range
    Dim startdate_2 As Range
    Dim completiondate_2 As Range
    Dim completiondate_3 As Range
    Dim Ticketnumber As Range
    Dim Workorder As Range
    Dim Workorder_2 As Range
    Dim Maintenancewindow As Range
    Dim OncallManager As Range
    Dim ProjectManager As Range
    Dim Projectmanagerinitials As Range
    Dim Projectmanagerphone As Range
    Dim footerSite As Range
    Dim footerSite_1 As Range

ActiveDocument.SaveAs2 FileName:=Environ("USERPROFILE") & "\Desktop\" & ProjectName.Text & " " & "Transfer Script" & " " & Month(Now) _
& "." & Day(Now) & "." & Year(Now) & ".docx"

With UserForm1
.Caption = "Initializing workbook..."
.Show vbModeless
.Repaint
End With

 'Inserting the UPS's to be transferred
    
'CreateObject("WScript.Shell").PopUp "Please wait while your document is created.", 1



If cboSite.value = "LAS 7" Then
LAS7transfer
Set UPS_1 = Nothing
Set UPS_2 = Nothing
Set UPS_3 = Nothing
Set UPS_4 = Nothing
Set UPS_5 = Nothing
Set UPS_6 = Nothing
Set UPS_7 = Nothing
Set UPS_8 = Nothing
Set UPS_9 = Nothing
Set UPS_10 = Nothing
Set UPS_11 = Nothing
Set UPS_12 = Nothing
Set Criticalitylevel = Nothing
Set buildingName = Nothing
Set address = Nothing
Set Projectmanagerphone = Nothing
Set Projectmanagerinitials = Nothing
Set ProjectManager = Nothing
Set OncallManager = Nothing
Set Maintenancewindow = Nothing
Set Workorder_2 = Nothing
Set Workorder = Nothing
Set Ticketnumber = Nothing
Set completiondate_3 = Nothing
Set completiondate_2 = Nothing
Set startdate_2 = Nothing
Set Startdate = Nothing
Set Endtime = Nothing
Set Starttime = Nothing

Else
    Set footerSite = ActiveDocument.Bookmarks("tsite").Range
    footerSite.Text = cboSite.value
    
    Set footerSite_1 = ActiveDocument.Bookmarks("tsite_1").Range
    footerSite_1.Text = cboSite.value
    
    Set Criticalitylevel = ActiveDocument.Bookmarks("tCriticalitylevel").Range
    Criticalitylevel.Text = cbocriticalitylevel.value
    
    Set buildingName = ActiveDocument.Bookmarks("tbuildingname").Range
    
    Set address = ActiveDocument.Bookmarks("taddress").Range
    
    Set Projectmanagerphone = ActiveDocument.Bookmarks("tProjectmanagerphone").Range
    Projectmanagerphone.Text = Me.Phonenumber.value
    
    Set Projectmanagerinitials = ActiveDocument.Bookmarks("tProjectManagerinitials").Range
    Projectmanagerinitials.Text = Me.Initials.value
    
    Set ProjectManager = ActiveDocument.Bookmarks("tProjectManager").Range
    ProjectManager.Text = Me.ProjectManager.value
    
    Set OncallManager = ActiveDocument.Bookmarks("tOncallmanager").Range
    OncallManager.Text = Me.oncall_1.value
    
    Set Maintenancewindow = ActiveDocument.Bookmarks("tMaintenancewindow").Range
    Maintenancewindow.Text = Me.Maintenancewindow.value
    
    Set Workorder_2 = ActiveDocument.Bookmarks("tworkorder_1").Range
    Workorder_2.Text = Me.Workorder_1.value
    
    Set Workorder = ActiveDocument.Bookmarks("tworkorder").Range
    Workorder.Text = Me.Workorder_1.value
    
    Set Ticketnumber = ActiveDocument.Bookmarks("tTicketnumber").Range
    Ticketnumber.Text = Me.ticketnumber_1.value
    
    Set completiondate_3 = ActiveDocument.Bookmarks("tcompletiondate_1").Range
    completiondate_3.Text = Me.completiondate_1.value
    
    Set completiondate_2 = ActiveDocument.Bookmarks("tcompletiondate").Range
    completiondate_2.Text = Me.completiondate_1.value
    
    Set startdate_2 = ActiveDocument.Bookmarks("tstartdate_1").Range
    startdate_2.Text = Me.startdate_1.value
    
    Set Startdate = ActiveDocument.Bookmarks("tStartdate").Range
    Startdate.Text = Me.startdate_1.value
    
    Set Endtime = ActiveDocument.Bookmarks("tendtime").Range
    Endtime.Text = Me.endtime_1.value
    
    Set Starttime = ActiveDocument.Bookmarks("tstarttime").Range
    Starttime.Text = Me.starttime_1.value

End If




'Selection of the table to be used


Set targetdoc = Documents.Open(FileName:=Environ("USERPROFILE") & "\Desktop\" & ProjectName.Text & " " & "transfer script" & " " & Month(Now) _
    & "." & Day(Now) & "." & Year(Now) & ".docx")
    

If (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Annual PM w/o Cal or Depletion" Then
            On Err.Number = 5174 GoTo Errorhandler
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Single UPS Annual or Corrective.docx")
            Set targettable = referencedoc.Tables(1)
            For Each targettable In referencedoc.Tables
               targettable.Range.Select
               Debug.Print targettable.Title
               Selection.Copy
               referencedoc.Close
               targetdoc.Activate
               Set startpoint = targetdoc.Paragraphs(146).Range
               startpoint.Paste
             Next targettable
             referencedoc.Close
            Set UPS_3 = ActiveDocument.Bookmarks("tUPS_3").Range
            Set UPS_4 = ActiveDocument.Bookmarks("tUPS_4").Range
            Set EquipmentID_2 = Nothing
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            setmbs
            setISX
            generatorselection_1
            
Errorhandler:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Single UPS Annual or Corrective.docx")
            End Select
Resume Next

    ElseIf (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
    cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Corrective Maintenance" Then
            On Err.Number = 5174 GoTo Errorhandler_1
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Single UPS Annual or Corrective.docx")
            Set targettable = referencedoc.Tables(1)
            For Each targettable In referencedoc.Tables
               targettable.Range.Select
               Debug.Print targettable.Title
               Selection.Copy
               referencedoc.Close
               targetdoc.Activate
               Set startpoint = targetdoc.Paragraphs(146).Range
               startpoint.Paste
             Next targettable
             referencedoc.Close
            Set Title = ActiveDocument.Bookmarks("ttitle").Range
            Set UPS = Nothing
            Set UPS_1 = Nothing
            Set UPS_2 = Nothing
            Set UPS_3 = ActiveDocument.Bookmarks("tUPS_3").Range
            Set UPS_4 = ActiveDocument.Bookmarks("tUPS_4").Range
            Set UPS_5 = Nothing
            Set UPS_6 = Nothing
            Set UPS_7 = Nothing
            Set EquipmentID_2 = Nothing
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            setmbs
            setISX
            generatorselection_1
            
Errorhandler_1:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Single UPS Annual or Corrective.docx")
            End Select
Resume Next

ElseIf (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal" Then
            On Err.Number = 5174 GoTo Errorhandler_2
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Single UPS Annual w Cal.docx")
            Set targettable = referencedoc.Tables(1)
             For Each targettable In referencedoc.Tables
                targettable.Range.Select
                Debug.Print targettable.Title
                Selection.Copy
                referencedoc.Close
                targetdoc.Activate
                Set startpoint = targetdoc.Paragraphs(146).Range
                startpoint.Paste
              Next targettable
              referencedoc.Close
            Set UPS_2 = ActiveDocument.Bookmarks("tUPS_2").Range
            Set UPS_3 = ActiveDocument.Bookmarks("tUPS_3").Range
            Set UPS_4 = ActiveDocument.Bookmarks("tUPS_4").Range
            Set UPS = ActiveDocument.Bookmarks("tUPS").Range
            Set UPS_1 = ActiveDocument.Bookmarks("tUPS_1").Range
            Set EquipmentID_2 = Nothing
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            Set UPS_5 = Nothing
            Set UPS_6 = Nothing
            Set UPS_7 = Nothing
            Set UPS_8 = Nothing
            Set UPS_9 = Nothing
            Set UPS_10 = Nothing
            Set UPS_11 = Nothing
            Set UPS_12 = Nothing
            Set UPS_2 = Nothing
            UPS.Text = cboEquipmentID.value
            UPS_1.Text = cboEquipmentID.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            setmbs
            setISX
            generatorSelection

Errorhandler_2:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Single UPS Annual w Cal.docx")
            End Select
Resume Next

ElseIf (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
cbonumberofups.value = "1" And cbotypeofmaintenance.value = "1 UPS Annual PM w/ Cal and Depletion" Then
            On Err.Number = 5174 GoTo Errorhandler_3
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Single UPS Annual w Cal and Depl.docx")
            Set targettable = referencedoc.Tables(1)
             For Each targettable In referencedoc.Tables
                targettable.Range.Select
                Debug.Print targettable.Title
                Selection.Copy
                referencedoc.Close
                targetdoc.Activate
                Set startpoint = targetdoc.Paragraphs(146).Range
                startpoint.Paste
              Next targettable
              referencedoc.Close
            Set UPS = ActiveDocument.Bookmarks("tUPS").Range
            Set UPS_1 = ActiveDocument.Bookmarks("tUPS_1").Range
            Set EquipmentID_2 = Nothing
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            Set UPS_5 = Nothing
            Set UPS_6 = Nothing
            Set UPS_7 = Nothing
            Set UPS_8 = Nothing
            Set UPS_9 = Nothing
            Set UPS_10 = Nothing
            Set UPS_11 = Nothing
            Set UPS_12 = Nothing
            Set UPS_2 = Nothing
            UPS.Text = cboEquipmentID.value
            UPS_1.Text = cboEquipmentID.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            setmbs
            setISX
            generatorSelection
        
Errorhandler_3:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Single UPS Annual w Cal and Depl.docx")
            End Select
Resume Next


ElseIf (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Annual PM w/o Cal or Depletion" Then
            On Err.Number = 5174 GoTo Errorhandler_4
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Multiple UPS Annual or Corrective.docx")
            Set targettable = referencedoc.Tables(1)
             For Each targettable In referencedoc.Tables
                targettable.Range.Select
                Debug.Print targettable.Title
                Selection.Copy
                referencedoc.Close
                targetdoc.Activate
                Set startpoint = targetdoc.Paragraphs(146).Range
                startpoint.Paste
             Next targettable
              referencedoc.Close
            Set Title = ActiveDocument.Bookmarks("ttitle").Range
            Set UPS = ActiveDocument.Bookmarks("tUPS").Range
            Set UPS_1 = ActiveDocument.Bookmarks("tUPS_1").Range
            Set UPS_2 = ActiveDocument.Bookmarks("tUPS_2").Range
            Set UPS_3 = ActiveDocument.Bookmarks("tUPS_3").Range
            Set UPS_4 = ActiveDocument.Bookmarks("tUPS_4").Range
            Set UPS_5 = ActiveDocument.Bookmarks("tUPS_5").Range
            Set UPS_6 = ActiveDocument.Bookmarks("tUPS_6").Range
            Set UPS_7 = ActiveDocument.Bookmarks("tUPS_7").Range
            Set UPS_8 = ActiveDocument.Bookmarks("tUPS_8").Range
            Set UPS_9 = Nothing
            Set UPS_10 = Nothing
            Set UPS_11 = Nothing
            Set UPS_12 = Nothing
            Set EquipmentID_2 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
            EquipmentID_2.Text = "And" & cboEquipmentID_2.value
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            UPS.Text = cboEquipmentID.value
            UPS_1.Text = cboEquipmentID.value
            UPS_2.Text = cboEquipmentID_2.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            UPS_5.Text = cboEquipmentID_2.value
            UPS_6.Text = cboEquipmentID.value
            UPS_7.Text = cboEquipmentID_2.value
            UPS_8.Text = cboEquipmentID_2.value
            setmbs
            setISX
            generatorselection_1
            
Errorhandler_4:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Multiple UPS Annual or Corrective.docx")
            End Select
Resume Next

ElseIf (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Corrective Maintenance" Then
            On Err.Number = 5174 GoTo Errorhandler_5
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Multiple UPS Annual or Corrective.docx")
            Set targettable = referencedoc.Tables(1)
             For Each targettable In referencedoc.Tables
                targettable.Range.Select
                Debug.Print targettable.Title
                Selection.Copy
                referencedoc.Close
                targetdoc.Activate
                Set startpoint = targetdoc.Paragraphs(146).Range
                startpoint.Paste
              Next targettable
              referencedoc.Close
            Set UPS = ActiveDocument.Bookmarks("tUPS").Range
            Set UPS_1 = ActiveDocument.Bookmarks("tUPS_1").Range
            Set UPS_2 = ActiveDocument.Bookmarks("tUPS_2").Range
            Set UPS_3 = ActiveDocument.Bookmarks("tUPS_3").Range
            Set UPS_4 = ActiveDocument.Bookmarks("tUPS_4").Range
            Set UPS_5 = ActiveDocument.Bookmarks("tUPS_5").Range
            Set UPS_6 = ActiveDocument.Bookmarks("tUPS_6").Range
            Set UPS_7 = ActiveDocument.Bookmarks("tUPS_7").Range
            Set UPS_8 = ActiveDocument.Bookmarks("tUPS_8").Range
            Set UPS_9 = Nothing
            Set UPS_10 = Nothing
            Set UPS_11 = Nothing
            Set UPS_12 = Nothing
            Set MBS_GSB = ActiveDocument.Bookmarks("tMBS_GSB").Range
            Set MVS = ActiveDocument.Bookmarks("tMVS").Range
            Set EquipmentID_2 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
            EquipmentID_2.Text = "And" & cboEquipmentID_2.value
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            UPS.Text = cboEquipmentID.value
            UPS_1.Text = cboEquipmentID.value
            UPS_2.Text = cboEquipmentID_2.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            UPS_5.Text = cboEquipmentID_2.value
            UPS_6.Text = cboEquipmentID.value
            UPS_7.Text = cboEquipmentID_2.value
            UPS_8.Text = cboEquipmentID_2.value
            setmbs
            setISX
            generatorselection_1
            
Errorhandler_5:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Multiple UPS Annual or Corrective.docx")
            End Select
Resume Next

ElseIf (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal" Then
            On Err.Number = 5174 GoTo Errorhandler_6
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Multiple UPS Annual w Cal.docx")
            Set targettable = referencedoc.Tables(1)
                For Each targettable In referencedoc.Tables
                   targettable.Range.Select
                   Debug.Print targettable.Title
                   Selection.Copy
                   referencedoc.Close
                   targetdoc.Activate
                   Set startpoint = targetdoc.Paragraphs(146).Range
                   startpoint.Paste
             Next targettable
              referencedoc.Close
            Set Title = ActiveDocument.Bookmarks("ttitle").Range
            Set UPS = ActiveDocument.Bookmarks("tUPS").Range
            Set UPS_1 = ActiveDocument.Bookmarks("tUPS_1").Range
            Set UPS_2 = ActiveDocument.Bookmarks("tUPS_2").Range
            Set UPS_3 = ActiveDocument.Bookmarks("tUPS_3").Range
            Set UPS_4 = ActiveDocument.Bookmarks("tUPS_4").Range
            Set UPS_5 = ActiveDocument.Bookmarks("tUPS_5").Range
            Set UPS_6 = ActiveDocument.Bookmarks("tUPS_6").Range
            Set UPS_7 = ActiveDocument.Bookmarks("tUPS_7").Range
            Set UPS_8 = ActiveDocument.Bookmarks("tUPS_8").Range
            Set UPS_9 = ActiveDocument.Bookmarks("tUPS_9").Range
            Set UPS_10 = ActiveDocument.Bookmarks("tUPS_10").Range
            Set UPS_11 = ActiveDocument.Bookmarks("tUPS_11").Range
            Set UPS_12 = ActiveDocument.Bookmarks("tUPS_12").Range
            Set EquipmentID_2 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
            EquipmentID_2.Text = "And" & cboEquipmentID_2.value
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            UPS.Text = cboEquipmentID.value
            UPS_1.Text = cboEquipmentID.value
            UPS_2.Text = cboEquipmentID_2.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            UPS_5.Text = cboEquipmentID_2.value
            UPS_6.Text = cboEquipmentID.value
            UPS_7.Text = cboEquipmentID_2.value
            UPS_8.Text = cboEquipmentID_2.value
            UPS_9.Text = cboEquipmentID.value
            UPS_10.Text = cboEquipmentID_2.value
            UPS_11.Text = cboEquipmentID.value
            UPS_12.Text = cboEquipmentID_2.value
            setmbs
            setISX
            generatorSelection
            
Errorhandler_6:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Multiple UPS Annual w Cal.docx")
            End Select
Resume Next

ElseIf (cboSite.value = "LAS 8" Or cboSite.value = "LAS 9" Or cboSite.value = "LAS 10" Or cboSite.value = "LAS 11") And _
cbonumberofups.value = "2" And cbotypeofmaintenance.value = "2 UPS's Annual PM w/ Cal and Depletion" Then
            On Err.Number = 5174 GoTo Errorhandler_7
            Set referencedoc = Documents.Open("T:\3 - SITE FILES & EQUIPMENT\EQUIPMENT (MULTI-SITE)\Transfer table reference documents\New Multiple UPS Annual w Cal and Depl.docx")
            Set targettable = referencedoc.Tables(1)
             For Each targettable In referencedoc.Tables
                targettable.Range.Select
                Debug.Print targettable.Title
                Selection.Copy
                referencedoc.Close
                targetdoc.Activate
                Set startpoint = targetdoc.Paragraphs(146).Range
                startpoint.Paste
              Next targettable
              referencedoc.Close
            Set UPS = ActiveDocument.Bookmarks("tUPS").Range
            Set UPS_1 = ActiveDocument.Bookmarks("tUPS_1").Range
            Set UPS_2 = ActiveDocument.Bookmarks("tUPS_2").Range
            Set UPS_3 = ActiveDocument.Bookmarks("tUPS_3").Range
            Set UPS_4 = ActiveDocument.Bookmarks("tUPS_4").Range
            Set UPS_5 = ActiveDocument.Bookmarks("tUPS_5").Range
            Set UPS_6 = ActiveDocument.Bookmarks("tUPS_6").Range
            Set UPS_7 = ActiveDocument.Bookmarks("tUPS_7").Range
            Set UPS_8 = ActiveDocument.Bookmarks("tUPS_8").Range
            Set UPS_9 = ActiveDocument.Bookmarks("tUPS_9").Range
            Set UPS_10 = ActiveDocument.Bookmarks("tUPS_10").Range
            Set UPS_11 = ActiveDocument.Bookmarks("tUPS_11").Range
            Set UPS_12 = ActiveDocument.Bookmarks("tUPS_12").Range
            Set EquipmentID_2 = ActiveDocument.Bookmarks("tEquipmentID_2").Range
            EquipmentID_2.Text = "And" & cboEquipmentID_2.value
            Set EquipmentID = ActiveDocument.Bookmarks("tEquipmentID").Range
            EquipmentID.Text = cboEquipmentID.value
            UPS.Text = cboEquipmentID.value
            UPS_1.Text = cboEquipmentID.value
            UPS_2.Text = cboEquipmentID_2.value
            UPS_3.Text = cboEquipmentID.value
            UPS_4.Text = cboEquipmentID.value
            UPS_5.Text = cboEquipmentID_2.value
            UPS_6.Text = cboEquipmentID.value
            UPS_7.Text = cboEquipmentID_2.value
            UPS_8.Text = cboEquipmentID_2.value
            UPS_9.Text = cboEquipmentID.value
            UPS_10.Text = cboEquipmentID_2.value
            UPS_11.Text = cboEquipmentID.value
            UPS_12.Text = cboEquipmentID_2.value
            setmbs
            setISX
            generatorSelection
            
Errorhandler_7:
            Select Case Err.Number
            Case 5174
            Err.Clear
            Set referencedoc = Documents.Open("https://intranet.switchnet.nv/CiOps/LAS7/Shared Documents/Transfer table reference documents/New Multiple UPS Annual w Cal and Depl.docx")
            End Select
Resume Next

End If


targetdoc.Activate

'number_UPS_Selection

If cbotypeofmaintenance.value = "" Then
    MsgBox "Please Select Type of Maintenance", vbOKCancel
    If vbOK = 1 Then
        Clear_Click
    Else
        Application.Quit
    End If
End If
        
    If cboPS.value = "" Then
        MsgBox "You must select a Power System", vbOKCancel
        If vbOK Then
            Clear_Click
        Else
            Application.Quit
        End If
    End If
    
      
    'Building Address and Name
    If cboSite.value = "LAS 8" Then
        address.Text = "5225 W Capovilla Ave," & vbCr & "Las Vegas, NV 89118"
        buildingName.Text = "LAS 8"
        setpduLAS8
    ElseIf cboSite.value = "LAS 9" Then
        address.Text = "7365 S Lindell Rd," & vbCr & "Las Vegas, NV 89139"
        buildingName.Text = "LAS 9"
        setpduLAS9
    ElseIf cboSite.value = "LAS 10" Then
        address.Text = "7365 S Lindell Rd," & vbCr & "Las Vegas, NV 89139"
        buildingName.Text = "LAS 10"
        setpduLAS10
    ElseIf cboSite.value = "LAS 11" Then
        address.Text = "7384 S Lindell Rd," & vbCr & "Las Vegas, NV 89139"
        buildingName.Text = "LAS 11"
    End If
    
    UserForm1.Hide
    MsgBox "This Transfer Script has been saved to your desktop"
    Me.Repaint
    Application.ScreenUpdating = True
    ClearClipboard_1
    
End Sub



