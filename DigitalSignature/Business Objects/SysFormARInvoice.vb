
Imports System.Drawing
Imports SAPbobsCOM

Namespace DigitalSignature
    Public Class SysFormARInvoice
        Public Const formtype As String = "133"
        Dim objForm, objUDFFormID As SAPbouiCOM.Form
        Dim strSQL As String
        Dim objRS As SAPbobsCOM.Recordset
        Dim oButton As SAPbouiCOM.Button
        Dim oEdit As SAPbouiCOM.EditText

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                If pval.BeforeAction Then
                    objForm = objaddon.objapplication.Forms.Item(FormUID)
                    Select Case pval.EventType

                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pval.ItemUID = "bgendsc" And (objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                                Dim UDFFileName As String
                                Dim pdffile As String()
                                If Not GetEnabledRpttoPDF = "Y" Then
                                    If objForm.Items.Item("txtpdf").Specific.String = "" Then
                                        objaddon.objapplication.SetStatusBarMessage("Please choose .pdf file...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        objForm.Items.Item("txtpdf").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                        BubbleEvent = False : Exit Sub
                                    Else
                                        UDFFileName = objForm.Items.Item("txtpdf").Specific.String 'objUDFForm.Items.Item("U_PDFFile").Specific.String
                                        pdffile = UDFFileName.Split(New String() {"."}, StringSplitOptions.None)
                                        If pdffile(pdffile.Length - 1).ToUpper <> "pdf".ToUpper Then
                                            objForm.Items.Item("txtpdf").Specific.String = ""
                                            objaddon.objapplication.StatusBar.SetText("Please select .pdf file...  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objForm.Items.Item("txtpdf").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                            BubbleEvent = False : Exit Sub
                                        End If
                                    End If
                                End If
                            End If

                    End Select
                Else
                    Select Case pval.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pval.ItemUID = "bgendsc" Then
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                                Dim TranName, RPTName, Strquery, GetLayout As String
                                Dim objRS As SAPbobsCOM.Recordset
                                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                TranName = objForm.DataSources.DBDataSources.Item(0).TableName '"SI"
                                If GetEnabledRpttoPDF = "Y" Then
                                    If objaddon.HANA Then
                                        GetLayout = objaddon.objglobalmethods.getSingleValue("select case when count(*)>1 then 1 else 0 end from ""@MIPL_ODSC"" T0 join ""@MIPL_DSC1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_TranName""='" & TranName & "'")
                                    Else
                                        GetLayout = objaddon.objglobalmethods.getSingleValue("select case when count(*)>1 then 1 else 0 end from [@MIPL_ODSC] T0 join [@MIPL_DSC1] T1 on T0.Code=T1.Code where T1.U_TranName='" & TranName & "'")
                                    End If
                                    If GetLayout = "1" Then
                                        If objaddon.HANA Then
                                            Query = "Select ""U_LayoutName"" as ""Layout"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'"
                                        Else
                                            Query = "Select U_LayoutName as Layout from  [@MIPL_DSC1] where U_TranName='" & TranName & "'"
                                        End If
                                        FrmMultiSel = objaddon.objapplication.Forms.ActiveForm
                                        If Not objaddon.FormExist("MULSEL") Then
                                            Dim Multiselect As New FrmMultiForm
                                            Multiselect.Show()
                                        End If
                                    Else
                                        objaddon.objapplication.StatusBar.SetText("Looking for the Crystal Report Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        If objaddon.HANA Then
                                            RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 ""U_RPTFile"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "' and ""U_RPTFile"" is not null order by ""LineId"" desc")
                                        Else
                                            RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 U_RPTFile from [@MIPL_DSC1] where U_TranName='" & TranName & "' and U_RPTFile is not null order by LineId desc")
                                        End If
                                        If objaddon.HANA Then
                                            Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                                        Else
                                            Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                                        End If
                                        objRS.DoQuery(Strquery)
                                        If objRS.RecordCount > 0 Then
                                            objaddon.objDSC.Create_RPT_To_PDF(objForm, RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                                            'objaddon.objDSC.Create_RPT_To_PDF_Test(FormCount, RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                                        Else
                                            objaddon.objapplication.StatusBar.SetText("No Data found to login Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    objaddon.objDSC.Create_Digital_Signature_Without_RPT(objForm.Items.Item("txtpdf").Specific.String, TranName)
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            oButton = objForm.Items.Item("bgendsc").Specific
                            oEdit = objForm.Items.Item("txtpdf").Specific
                            If GetEnabledRpttoPDF = "Y" Then
                                oButton.Item.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                                oButton.Item.Top = objForm.Items.Item("10000330").Top
                                oEdit.Item.Visible = False
                            Else
                                oEdit.Item.Visible = True
                                oButton.Item.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                                oButton.Item.Top = objForm.Items.Item("10000330").Top
                                oEdit.Item.Left = oButton.Item.Left - oEdit.Item.Width - 5 ' 125
                                oEdit.Item.Top = objForm.Items.Item("10000330").Top 'Button0.Item.Top
                            End If


                    End Select
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objForm = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                If BusinessObjectInfo.BeforeAction Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            Create_Customize_Fields(objForm.UniqueID)
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD


                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            oButton = objForm.Items.Item("bgendsc").Specific
                            oEdit = objForm.Items.Item("txtpdf").Specific
                            oButton.Item.Enabled = True
                            If objaddon.HANA Then
                                GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select ""U_RptToPdf"" from ""@MIPL_ODSC"" ") 'where ifnull(""U_RptToPdf"",'N')='N'
                            Else
                                GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select U_RptToPdf from [@MIPL_ODSC] ") 'where isnull(U_RptToPdf,'N')='N'
                            End If
                            If GetEnabledRpttoPDF = "Y" Then
                                oButton.Item.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                                oButton.Item.Top = objForm.Items.Item("10000330").Top
                                oEdit.Item.Visible = False
                            Else
                                oEdit.Item.Visible = True
                                oButton.Item.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                                oButton.Item.Top = objForm.Items.Item("10000330").Top
                                oEdit.Item.Left = oButton.Item.Left - oEdit.Item.Width - 5 ' 125
                                oEdit.Item.Top = objForm.Items.Item("10000330").Top 'Button0.Item.Top
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD


                    End Select

                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Create_Customize_Fields(ByVal oFormUID As String)
            Try
                objForm = clsModule.objaddon.objapplication.Forms.Item(oFormUID)
                Try
                    If objForm.Items.Item("bgendsc").UniqueID = "bgendsc" Or objForm.Items.Item("txtpdf").UniqueID = "txtpdf" Then
                        Exit Sub
                    End If
                Catch ex As Exception

                End Try
                Dim oItem As SAPbouiCOM.Item
                oItem = objForm.Items.Add("bgendsc", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                oButton = oItem.Specific 'CType(oItem.Specific, SAPbouiCOM.Button)
                oButton.Caption = "Generate DSC"
                oItem.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                oItem.Top = objForm.Items.Item("2").Top
                oItem.Height = objForm.Items.Item("2").Height
                oItem.LinkTo = "10000330"
                Dim Fieldsize As Size = System.Windows.Forms.TextRenderer.MeasureText("Generate DSC", New Font("Arial", 12.0F))
                oItem.Width = Fieldsize.Width - 20
                objForm.Items.Item("bgendsc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("bgendsc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False)


                oItem = objForm.Items.Add("txtpdf", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                oItem.Left = objForm.Items.Item("bgendsc").Left - objForm.Items.Item("10000330").Width - 5
                oItem.Width = objForm.Items.Item("4").Width - 10
                oItem.Top = objForm.Items.Item("2").Top
                oItem.Height = objForm.Items.Item("4").Height
                oItem.LinkTo = "bgendsc"
                oEdit = oItem.Specific
                oEdit.DataBind.SetBound(True, "OINV", "U_PDFFile")
                'oEdit.Item.LinkTo = "bgendsc"
            Catch ex As Exception
            End Try
        End Sub

    End Class

End Namespace
