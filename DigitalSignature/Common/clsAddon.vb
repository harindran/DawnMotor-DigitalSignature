Imports SAPbouiCOM.Framework
Imports System.IO


Namespace DigitalSignature
    Public Class clsAddon
        Public WithEvents objapplication As SAPbouiCOM.Application
        Public objcompany As SAPbobsCOM.Company
        Dim objmenuevent As clsMenuEvent
        Dim objrightclickevent As clsRightClickEvent
        Public objglobalmethods As clsGlobalMethods
        Public objDSC As ClsDSC
        'Public oARInvoice As SysFormARInvoice
        Dim objform As SAPbouiCOM.Form
        Dim strsql As String = ""
        Dim objrs As SAPbobsCOM.Recordset
        Dim print_close As Boolean = False
        Public HANA As Boolean = False
        'Public HANA As Boolean = True
        Public HWKEY() As String = New String() {"X1211807750", "L1653539483", "N1647886019"}

        Public Sub Intialize(ByVal args() As String)
            Try
                Dim oapplication As Application
                If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
                objapplication = Application.SBO_Application
                If isValidLicense() Then
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objcompany = Application.SBO_Application.Company.GetDICompany()

                    Create_Objects() 'Object Creation Part
                    Create_DatabaseFields() 'UDF & UDO Creation Part
                    Menu() 'Menu Creation Part
                    Add_Authorizations() 'User Permissions

                    objapplication.StatusBar.SetText("Digital Signature Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    oapplication.Run()
                Else
                    objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'System.Windows.Forms.Application.Run()
            Catch ex As Exception
                objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Function isValidLicense() As Boolean
            Try
                Try
                    If objapplication.Forms.ActiveForm.TypeCount > 1 Then
                        For i As Integer = 0 To objapplication.Forms.ActiveForm.TypeCount - 1
                            objapplication.Forms.ActiveForm.Close()
                        Next
                    End If
                Catch ex As Exception
                End Try

                objapplication.Menus.Item("257").Activate()
                Dim CrrHWKEY As String = objapplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
                objapplication.Forms.ActiveForm.Close()

                For i As Integer = 0 To HWKEY.Length - 1
                    If HWKEY(i).Trim = CrrHWKEY.Trim Then
                        Return True
                    End If
                Next
                MsgBox("Installing Add-On failed due to License mismatch", MsgBoxStyle.OkOnly, "License Management")
                Return False
            Catch ex As Exception
                'MsgBox(ex.ToString)
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
            Return True
        End Function

        Private Sub Create_Objects()
            objmenuevent = New clsMenuEvent
            objrightclickevent = New clsRightClickEvent
            objglobalmethods = New clsGlobalMethods
            objDSC = New ClsDSC
            'oARInvoice = New SysFormARInvoice
        End Sub

        Private Sub Create_DatabaseFields()
            'If objapplication.MessageBox("Do you want to execute the field Creations?", 2, "Yes", "No") <> 1 Then Exit Sub
            
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Dim objtable As New clsTable
            objtable.FieldCreation()
            If objaddon.HANA Then
                GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select ""U_RptToPdf"" from ""@MIPL_ODSC"" ")
            Else
                GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select U_RptToPdf from [@MIPL_ODSC] ")
            End If
            'End If

        End Sub

        Public Sub Add_Authorizations()
            Try
                objaddon.objglobalmethods.AddToPermissionTree("Altrocks Tech", "ATPL_ADD-ON", "", "", "Y"c) 'Level 1 - Company Name

                objaddon.objglobalmethods.AddToPermissionTree("Digital Signature", "ATPL_DSC", "", "ATPL_ADD-ON", "Y"c) 'Level 2 - Add-on Name

                objaddon.objglobalmethods.AddToPermissionTree("DSC Settings", "ATPL_CFGSET", "MIDSC", "ATPL_DSC", "Y"c) 'SubLevel of Level 2 - Screen Name

            Catch ex As Exception
            End Try
        End Sub

#Region "Menu Creation Details"

        Private Sub Menu()
            Dim Menucount As Integer = 11
            CreateMenu("", Menucount, "Digital Signature", SAPbouiCOM.BoMenuType.mt_POPUP, "DigSign", "43525")
            Menucount = 1 'Menu Inside            
            CreateMenu("", Menucount, "Digital Signature", SAPbouiCOM.BoMenuType.mt_STRING, "DSC", "DigSign") : Menucount += 1

        End Sub

        Private Sub CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenuID As String)
            Try
                Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
                Dim parentmenu As SAPbouiCOM.MenuItem
                parentmenu = objapplication.Menus.Item(ParentMenuID)
                If parentmenu.SubMenus.Exists(UniqueID.ToString) Then Exit Sub
                oMenuPackage = objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oMenuPackage.Image = ImagePath
                oMenuPackage.Position = Position
                oMenuPackage.Type = MenuType
                oMenuPackage.UniqueID = UniqueID
                oMenuPackage.String = DisplayName
                parentmenu.SubMenus.AddEx(oMenuPackage)
            Catch ex As Exception
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End Try
            'Return ParentMenu.SubMenus.Item(UniqueID)
        End Sub

#End Region

#Region "ItemEvent_Link Button"

        Private Sub objapplication_ItemEvent(FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objapplication.ItemEvent
            Try
                'Select Case pVal.FormTypeEx
                '    Case SysFormARInvoice.formtype
                '        oARInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                'End Select

                objform = objaddon.objapplication.Forms.Item(FormUID)
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            'Dim oform = objapplication.Forms.ActiveForm
                            If bModal And objaddon.objapplication.Forms.ActiveForm.TypeEx <> "MULSEL" Then
                                BubbleEvent = False
                                objapplication.Forms.Item("MULSEL").Select() : Exit Sub
                            End If
                            If pVal.ItemUID = "bgendsc" Then
                                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then BubbleEvent = False : Exit Sub
                                Dim UDFFileName As String
                                Dim pdffile As String()
                                If Not GetEnabledRpttoPDF = "Y" Then
                                    If objform.Items.Item("txtpdf").Specific.String = "" Then
                                        objaddon.objapplication.SetStatusBarMessage("Please choose .pdf file...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                                        objform.Items.Item("txtpdf").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                        BubbleEvent = False : Exit Sub
                                    Else
                                        UDFFileName = objform.Items.Item("txtpdf").Specific.String 'objUDFForm.Items.Item("U_PDFFile").Specific.String
                                        pdffile = UDFFileName.Split(New String() {"."}, StringSplitOptions.None)
                                        If pdffile(pdffile.Length - 1).ToUpper <> "pdf".ToUpper Then
                                            objform.Items.Item("txtpdf").Specific.String = ""
                                            objaddon.objapplication.StatusBar.SetText("Please select .pdf file...  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objform.Items.Item("txtpdf").Click(SAPbouiCOM.BoCellClickType.ct_Double)
                                            BubbleEvent = False : Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            Dim EventEnum As SAPbouiCOM.BoEventTypes
                            EventEnum = pVal.EventType
                            If FormUID = "MULSEL" And (EventEnum = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE) And bModal Then
                                bModal = False
                            End If

                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                            objDSC.Resize_Customize_Fields(objform)
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "bgendsc" Then
                                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                                Dim TranName, RPTName, Strquery, GetLayout As String
                                Dim objRS As SAPbobsCOM.Recordset
                                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                TranName = objform.DataSources.DBDataSources.Item(0).TableName '"SI"
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
                                            objaddon.objDSC.Create_RPT_To_PDF(objform, RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                                        Else
                                            objaddon.objapplication.StatusBar.SetText("No Data found to login Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If
                                    End If
                                Else
                                    objaddon.objDSC.Create_Digital_Signature_Without_RPT(objform.Items.Item("txtpdf").Specific.String, TranName)
                                End If
                            End If
                    End Select
                End If

            Catch ex As Exception

            End Try
        End Sub

#End Region

#Region "FormDataEvent"

        Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objapplication.FormDataEvent
            Try
                'Select Case BusinessObjectInfo.FormTypeEx
                '    Case SysFormARInvoice.formtype
                '        oARInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                'End Select
                objform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                If BusinessObjectInfo.BeforeAction Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            If HANA Then
                                strsql = objglobalmethods.getSingleValue("Select 1 as Stat from ""@MIPL_DSC1"" T0 Left Join ""@MIPL_ODSC"" T1 On T0.""Code""=T1.""Code""  Where T0.""U_RPTFile"" is not null and T0.""U_TranName""='" & objform.DataSources.DBDataSources.Item(0).TableName & "'")
                            Else
                                strsql = objglobalmethods.getSingleValue("Select 1 as Stat from [@MIPL_DSC1] T0 Left Join [@MIPL_ODSC] T1 On T0.Code=T1.Code  Where T0.U_RPTFile is not null and T0.U_TranName='" & objform.DataSources.DBDataSources.Item(0).TableName & "'")
                            End If
                            If strsql = "1" Then objDSC.Create_Customize_Fields(objform)
                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            objDSC.Resize_Customize_Fields(objform)
                    End Select
                End If

            Catch ex As Exception

            End Try

        End Sub

#End Region

#Region "Menu Event"

        Public Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objapplication.MenuEvent
            Try
                Select Case pVal.MenuUID
                    Case "1281", "1282", "1283", "1284", "1285", "1286", "1287", "1300", "1288", "1289", "1290", "1291", "1304", "1292", "1293", "MIDSC"
                        objmenuevent.MenuEvent_For_StandardMenu(pVal, BubbleEvent)
                    Case "DSC"
                        MenuEvent_For_FormOpening(pVal, BubbleEvent)
                        'Case "1293"
                        '    BubbleEvent = False
                    Case "519"
                        MenuEvent_For_Preview(pVal, BubbleEvent)
                End Select
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in SBO_Application MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Public Sub MenuEvent_For_Preview(ByRef pval As SAPbouiCOM.MenuEvent, ByRef bubbleevent As Boolean)
            Dim oform = objapplication.Forms.ActiveForm()
            'If pval.BeforeAction Then
            '    If oform.TypeEx = "TRANOLVA" Then MenuEvent_For_PrintPreview(oform, "8f481d5cf08e494f9a83e1e46ab2299e", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "TRANOLAP" Then MenuEvent_For_PrintPreview(oform, "f15ee526ac514070a9d546cda7f94daf", "txtentry") : bubbleevent = False
            '    If oform.TypeEx = "OLSE" Then MenuEvent_For_PrintPreview(oform, "e47ed373e0cc48efb47c9773fba64fc3", "txtentry") : bubbleevent = False
            'End If
        End Sub

        Private Sub MenuEvent_For_PrintPreview(ByVal oform As SAPbouiCOM.Form, ByVal Menuid As String, ByVal Docentry_field As String)
            Try
                Dim Docentry_Est As String = oform.Items.Item(Docentry_field).Specific.String
                If Docentry_Est = "" Then Exit Sub
                print_close = False
                objapplication.Menus.Item(Menuid).Activate()
                oform = objapplication.Forms.ActiveForm()
                oform.Items.Item("1000003").Specific.string = Docentry_Est
                oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                print_close = True
            Catch ex As Exception
            End Try
        End Sub

        Public Sub MenuEvent_For_FormOpening(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                If pVal.BeforeAction = False Then
                    Select Case pVal.MenuUID
                        Case "DSC"
                            If Not FormExist("MIDSC") Then
                                Dim activeform As New FrmDSCSettings
                                activeform.Show()
                            End If

                    End Select

                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Form Opening MenuEvent" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Public Function FormExist(ByVal FormID As String) As Boolean
            FormExist = False
            For Each uid As SAPbouiCOM.Form In objaddon.objapplication.Forms
                If uid.UniqueID = FormID Then
                    FormExist = True
                    Exit For
                End If
            Next
            If FormExist Then
                If FormID = "MULSEL" Then
                    Try
                        Dim cflForm As SAPbouiCOM.Form
                        If objaddon.objapplication.Forms.Count > 0 Then
                            For frm As Integer = 0 To objaddon.objapplication.Forms.Count - 1
                                If objaddon.objapplication.Forms.Item(frm).UniqueID = "MULSEL" Then
                                    cflForm = objaddon.objapplication.Forms.Item("MULSEL")
                                    cflForm.Close()
                                    Return False
                                    Exit For
                                End If
                            Next
                        End If
                    Catch ex As Exception
                    End Try
                Else
                    objaddon.objapplication.Forms.Item(FormID).Visible = True
                    objaddon.objapplication.Forms.Item(FormID).Select()
                End If
            End If
        End Function

#End Region



#Region "LayoutKeyEvent"

        Public Sub SBO_Application_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles objapplication.LayoutKeyEvent
            'Dim oForm_Layout As SAPbouiCOM.Form = Nothing
            'If SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.BusinessObject.Type = "NJT_CES" Then
            '    oForm_Layout = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(eventInfo.FormUID)
            'End If
        End Sub

#End Region

#Region "Application Event"

        Public Sub SBO_Application_AppEvent(EventType As SAPbouiCOM.BoAppEventTypes) Handles objapplication.AppEvent
            If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
                Try
                    Remove_Menu({"43525,DigSign"})
                    DisConnect_Addon()
                Catch ex As Exception

                End Try
            End If

        End Sub

        Private Sub DisConnect_Addon()
            Try
                If objaddon.objapplication.Forms.Count > 0 Then
                    Try
                        For frm As Integer = objaddon.objapplication.Forms.Count - 1 To 0 Step -1
                            If objaddon.objapplication.Forms.Item(frm).IsSystem = True Then Continue For
                            objaddon.objapplication.Forms.Item(frm).Close()
                        Next
                    Catch ex As Exception
                    End Try

                    'If objApplication.Menus.Item("43520").SubMenus.Exists(MenuID) Then objApplication.Menus.Item("43520").SubMenus.RemoveEx(MenuID)
                End If
                If objcompany.Connected Then objcompany.Disconnect()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication)
                objcompany = Nothing
                GC.Collect()
                System.Windows.Forms.Application.Exit()
                End
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Remove_Menu(ByVal MenuID() As String)
            Try
                Dim split_char() As String

                If Not MenuID Is Nothing Then
                    If MenuID.Length > 0 Then
                        For i = 0 To MenuID.Length - 1
                            If Trim(MenuID(i)) = "" Then Continue For
                            split_char = MenuID(i).Split(",")
                            If split_char.Length <> 2 Then Continue For
                            If (objaddon.objapplication.Menus.Item(split_char(0)).SubMenus.Exists(split_char(1))) Then
                                objaddon.objapplication.Menus.Item(split_char(0)).SubMenus.RemoveEx(split_char(1))
                            End If
                        Next
                    End If
                End If



            Catch ex As Exception

            End Try
        End Sub

#End Region

#Region "Right Click Event"

        Private Sub objapplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objapplication.RightClickEvent
            Try
                Select Case objapplication.Forms.ActiveForm.TypeEx
                    Case "MIDSC"
                        objrightclickevent.RightClickEvent(eventInfo, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

#End Region


    End Class
End Namespace
