Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System
Imports System.Text
Imports System.Threading.Tasks
Imports System.Xml
Imports System.Linq
Imports System.IO
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports iTextSharp.text.pdf
Imports CrystalDecisions.CrystalReports.Engine
Imports BcX509 = Org.BouncyCastle.X509
Imports Org.BouncyCastle.Pkcs
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.X509
'Imports DotNetUtils = Org.BouncyCastle.Security.DotNetUtils
Imports System.Security.Cryptography.X509Certificates
Imports SAPbobsCOM
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports
Imports iTextSharp.text
Imports iTextSharp.text.pdf.parser
Imports DigitalSignature.ClsPDFText
Imports iTextSharp.text.pdf.security

Namespace DigitalSignature
    Public Class ClsDSC

        Dim oButton As SAPbouiCOM.Button
        Dim oEdit As SAPbouiCOM.EditText

        Public Sub Create_RPT_To_PDF(ByVal objDSCForm As SAPbouiCOM.Form, ByVal RPTFileName As String, ByVal ServerName As String, ByVal DBName As String, ByVal DBUserName As String, ByVal DbPassword As String, ByVal TranName As String, Optional LayoutName As String = "")
            Try
                Dim cryRpt As New ReportDocument
                Dim rName, SavePDFFile, Foldername, DSCPDFFile, SysPath As String
                Dim Strquery, ParamQuery As String, ParamValue As String = ""
                Dim objRS, objRSparam As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRSparam = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                If objaddon.HANA Then
                    Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                Else
                    Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                End If
                objRS.DoQuery(Strquery)
                cryRpt.Load(RPTFileName)
                cryRpt.DataSourceConnections(0).SetConnection(ServerName, DBName, False)
                cryRpt.DataSourceConnections(0).SetLogon(DBUserName, DbPassword)
                Try
                    cryRpt.Refresh()
                    cryRpt.VerifyDatabase()
                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText("Verify Database: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
                If objaddon.HANA Then
                    ParamQuery = "Select ""U_ParamName"",""U_ParamVal"",""U_Textinpdf"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'"
                    If LayoutName <> "" Then ParamQuery += vbCrLf + " and ""U_LayoutName""='" & LayoutName & "' "
                Else
                    ParamQuery = "Select U_ParamName,U_ParamVal,U_Textinpdf from [@MIPL_DSC1] where U_TranName='" & TranName & "'"
                    If LayoutName <> "" Then ParamQuery += vbCrLf + " and U_LayoutName='" & LayoutName & "' "
                End If
                objRSparam.DoQuery(ParamQuery)
                'Dim objDSCForm As SAPbouiCOM.Form = Nothing
                'Dim Theader As String = ""
                'If TranName = "SI" Then
                '    Theader = "OINV"
                '    objDSCForm = objaddon.objapplication.Forms.GetForm("133", TypeCount)
                'ElseIf TranName = "DC" Then
                '    Theader = "ODLN"
                '    objDSCForm = objaddon.objapplication.Forms.GetForm("140", TypeCount)
                'ElseIf TranName = "SR" Then
                '    Theader = "ORIN"
                '    objDSCForm = objaddon.objapplication.Forms.GetForm("179", TypeCount)
                'ElseIf TranName = "PO" Then
                '    Theader = "OPOR"
                '    objDSCForm = objaddon.objapplication.Forms.GetForm("142", TypeCount)
                'End If
                If objRSparam.Fields.Item("U_ParamVal").Value.ToString.ToUpper = "DocEntry".ToUpper Then
                    ParamValue = objDSCForm.DataSources.DBDataSources.Item(TranName).GetValue("DocEntry", 0) 'objaddon.objapplication.Forms.ActiveForm.DataSources.DBDataSources.Item(Theader).GetValue("DocEntry", 0)
                End If
                If ParamValue = "" Then
                    objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + "Unable to get the DocEntry please re-open the transaction screen...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'objaddon.objglobalmethods.WriteErrorLog("FileName: " + RPTFileName + " ParamVal: " + ParamValue + " TableName: " + Theader)
                cryRpt.SetParameterValue(Trim(objRSparam.Fields.Item("U_ParamName").Value.ToString), CStr(ParamValue))

                rName = SystemInformation.UserName
                'objaddon.objglobalmethods.WriteErrorLog("UserName" + rName)
                If objaddon.HANA Then
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select ""AttachPath"" from OADP")
                Else
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select AttachPath from OADP")
                End If
                'objaddon.objglobalmethods.WriteErrorLog("SysPath" + SysPath)
                'Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF"
                Foldername = SysPath + rName + "\" + objaddon.objcompany.UserName + "\PDF"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                SavePDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                'objaddon.objglobalmethods.WriteErrorLog("SavePDFFile" + SavePDFFile)
                If File.Exists(SavePDFFile) Then
                    File.Delete(SavePDFFile)
                End If
                cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, SavePDFFile)
                'objaddon.objglobalmethods.WriteErrorLog("ExportToDisk")
                'cryRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, SavePDFFile)
                cryRpt.Close()
                'objaddon.objglobalmethods.WriteErrorLog("cryRpt Close")
                'Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF\DSC"
                Foldername = SysPath + rName + "\" + objaddon.objcompany.UserName + "\PDF\DSC"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                DSCPDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(DSCPDFFile) Then
                    File.Delete(DSCPDFFile)
                End If
                'objaddon.objDSC.AddSignNameinPDF(SavePDFFile, DSCPDFFile)
                Create_Digital_Signature(objRS.Fields.Item("U_PFXFile").Value, objRS.Fields.Item("U_PFXPass").Value, SavePDFFile, DSCPDFFile, TranName, LayoutName)
                objRS = Nothing
                objRSparam = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Sub Create_RPT_To_PDF_Test(ByVal TypeCount As Integer, RPTFileName As String, ByVal ServerName As String, ByVal DBName As String, ByVal DBUserName As String, ByVal DbPassword As String, ByVal TranName As String)
            Dim crzReport As New ReportDocument
            Dim sDocOutPath As String = Nothing
            Dim sCreatePDFDebug As String = Nothing
            Dim CrzPdfOptions As New PdfFormatOptions
            Dim CrzExportOptions As New ExportOptions
            Dim CrzDiskFileDestinationOptions As New DiskFileDestinationOptions()
            Dim CrzFormatTypeOptions As New PdfRtfWordFormatOptions()
            Try
                Dim rName, SavePDFFile, Foldername, DSCPDFFile, SysPath As String
                Dim Strquery, ParamQuery As String, ParamValue As String = ""
                Dim objRS, objRSparam As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRSparam = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objaddon.objapplication.StatusBar.SetText("Generating RPT to PDF Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If objaddon.HANA Then
                    Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                Else
                    Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                End If
                objRS.DoQuery(Strquery)

                If objaddon.HANA Then
                    ParamQuery = "Select ""U_ParamName"",""U_ParamVal"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'"
                Else
                    ParamQuery = "Select U_ParamName,U_ParamVal from [@MIPL_DSC1] where U_TranName='" & TranName & "'"
                End If
                objRSparam.DoQuery(ParamQuery)
                Dim objDSCForm As SAPbouiCOM.Form = Nothing
                Dim Theader As String = ""
                If TranName = "SI" Then
                    Theader = "OINV"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("133", TypeCount)
                ElseIf TranName = "DC" Then
                    Theader = "ODLN"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("140", TypeCount)
                ElseIf TranName = "SR" Then
                    Theader = "ORIN"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("179", TypeCount)
                ElseIf TranName = "PO" Then
                    Theader = "OPOR"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("142", TypeCount)
                End If
                If objRSparam.Fields.Item("U_ParamVal").Value.ToString.ToUpper = "DocEntry".ToUpper Then
                    ParamValue = objDSCForm.DataSources.DBDataSources.Item(Theader).GetValue("DocEntry", 0) 'objaddon.objapplication.Forms.ActiveForm.DataSources.DBDataSources.Item(Theader).GetValue("DocEntry", 0)
                End If
                If ParamValue = "" Then
                    objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + "Unable to get the docentry please re-open the transaction screen...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'objaddon.objglobalmethods.WriteErrorLog("FileName: " + RPTFileName + " ParamVal: " + ParamValue + " TableName: " + Theader)
                crzReport.Load(RPTFileName)
                Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                Dim crParameterFieldDefinition As ParameterFieldDefinition
                Dim crParameterValues As New ParameterValues
                Dim crParameterDiscreteValue As New ParameterDiscreteValue

                Dim crTable As Engine.Table
                Dim crTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
                Dim ConnInfo As New CrystalDecisions.Shared.ConnectionInfo
                ConnInfo.ServerName = ServerName
                ConnInfo.DatabaseName = DBName
                ConnInfo.UserID = DBUserName
                ConnInfo.Password = DbPassword

                For Each crTable In crzReport.Database.Tables
                    crTableLogonInfo = crTable.LogOnInfo
                    crTableLogonInfo.ConnectionInfo = ConnInfo
                    crTable.ApplyLogOnInfo(crTableLogonInfo)
                Next


                crParameterDiscreteValue.Value = ParamValue
                crParameterFieldDefinitions = crzReport.DataDefinition.ParameterFields()
                crParameterFieldDefinition = crParameterFieldDefinitions.Item(Trim(objRSparam.Fields.Item("U_ParamName").Value.ToString))
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                rName = SystemInformation.UserName
                If objaddon.HANA Then
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select ""AttachPath"" from OADP")
                Else
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select AttachPath from OADP")
                End If

                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                SavePDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(SavePDFFile) Then
                    File.Delete(SavePDFFile)
                End If
                'cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, SavePDFFile)
                'cryRpt.Close()
                CrzDiskFileDestinationOptions.DiskFileName = SavePDFFile 'Set the destination path and file name
                CrzExportOptions = crzReport.ExportOptions 'Set export options
                With CrzExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile ' DiskFile, ExchangeFolder, MicrosoftMail, NoDestination
                    .ExportFormatType = ExportFormatType.PortableDocFormat 'ExcelWorkBook, HTML32, HTML40, NoFormat, PDF, RichText, RTPR, TabSeperatedText, Text
                    .DestinationOptions = CrzDiskFileDestinationOptions
                    .FormatOptions = CrzFormatTypeOptions
                End With
                crzReport.Export()
                'crzReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, SavePDFFile)
                crParameterFieldDefinition.CurrentValues.Clear()
                objaddon.objapplication.StatusBar.SetText("PDF File generated Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF\DSC"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                DSCPDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(DSCPDFFile) Then
                    File.Delete(DSCPDFFile)
                End If
                'objaddon.objDSC.AddSignNameinPDF(SavePDFFile, DSCPDFFile)
                Create_Digital_Signature(objRS.Fields.Item("U_PFXFile").Value, objRS.Fields.Item("U_PFXPass").Value, SavePDFFile, DSCPDFFile)
                objRS = Nothing
                objRSparam = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                crzReport.Close()
                crzReport.Dispose()
                CrzPdfOptions = Nothing
                CrzExportOptions = Nothing
                CrzDiskFileDestinationOptions = Nothing
                CrzFormatTypeOptions = Nothing
            End Try
        End Sub

        Private Sub Create_Digital_Signature(ByVal PFXFile As String, ByVal PFXPassword As String, ByVal ReadPDF As String, ByVal FinalPDFwithDSC As String, Optional ByVal TranName As String = "", Optional LayoutName As String = "")
            Try
                objaddon.objapplication.StatusBar.SetText("Applying DSC Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim myCert As PDFSigner.Cert = Nothing
                Dim SignerName, StrQuery As String
                Dim signer As String()
                Dim DSCErrorflag As Boolean = False
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                myCert = New PDFSigner.Cert(PFXFile, PFXPassword)
                Dim pfxKeyStore As Pkcs12Store = New Pkcs12Store(New FileStream(PFXFile, FileMode.Open, FileAccess.Read), PFXPassword.ToCharArray())
                Dim collect As New X509Certificate2Collection
                collect.Import(PFXFile, PFXPassword, X509KeyStorageFlags.PersistKeySet)
                For Each cert In collect
                    SignerName = cert.Issuer
                Next
                signer = SignerName.Split(New String() {"CN="}, StringSplitOptions.None)
                SignerName = signer(1).ToString
                Dim Reader As New PdfReader(ReadPDF)
                Dim md As New PDFSigner.MetaData
                md.Info1 = Reader.Info

                Dim MyMD As New PDFSigner.MetaData
                'Dim pdfs As PDFSigner.PDFSigner = New PDFSigner.PDFSigner(ReadPDF, FinalPDFwithDSC, myCert, MyMD)
                If objaddon.HANA Then
                    StrQuery = "select ""U_TxtReason"",""U_TxtLocation"",""U_ValSym"" ""Valid Symbol"",""U_Format"" ""Format"" from ""@MIPL_ODSC"" "
                Else
                    StrQuery = "select U_TxtReason,U_TxtLocation,U_ValSym [Valid Symbol],U_Format Format from [@MIPL_ODSC]"
                End If
                objRs.DoQuery(StrQuery)

                Dim Reason, Location, GetText As String
                Dim llx, lly, urx, ury As Integer
                Dim position As Single()
                Dim x, y As Single
                Dim page As Integer
                'If objaddon.objapplication.Forms.ActiveForm.Type.ToString = "133" Then   'AR Invoice
                '    TranName = "SI"
                'ElseIf objaddon.objapplication.Forms.ActiveForm.Type.ToString = "140" Then  'Delivery
                '    TranName = "DC"
                'ElseIf objaddon.objapplication.Forms.ActiveForm.Type.ToString = "142" Then  'Purchase Order
                '    TranName = "PO"
                'ElseIf objaddon.objapplication.Forms.ActiveForm.Type.ToString = "179" Then  'AR Credit Memo
                '    TranName = "SR"
                'Else
                '    Exit Sub
                'End If
                If objaddon.HANA Then
                    StrQuery = "Select Top 1 ""U_Textinpdf"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'"
                    If LayoutName <> "" Then StrQuery += vbCrLf + " and ""U_LayoutName""='" & LayoutName & "' "
                    GetText = objaddon.objglobalmethods.getSingleValue(StrQuery)
                Else
                    StrQuery = "select Top 1 U_Textinpdf from [@MIPL_DSC1] where U_TranName='" & TranName & "'"
                    If LayoutName <> "" Then StrQuery += vbCrLf + " and U_LayoutName='" & LayoutName & "' "
                    GetText = objaddon.objglobalmethods.getSingleValue(StrQuery)
                End If
                If Trim(GetText) = "" Then
                    objaddon.objapplication.StatusBar.SetText("Layout: " + LayoutName + " Text is not defined in DSC Settings...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If

                If objaddon.HANA Then
                    Reason = objaddon.objglobalmethods.getSingleValue("select ""U_Reason"" from ""@MIPL_ODSC""")
                    Location = objaddon.objglobalmethods.getSingleValue("select ""U_Location"" from ""@MIPL_ODSC""")
                Else
                    Reason = objaddon.objglobalmethods.getSingleValue("select U_Reason from [@MIPL_ODSC]")
                    Location = objaddon.objglobalmethods.getSingleValue("select U_Location from [@MIPL_ODSC]")
                End If
                Dim ErrFlagCount As Integer = 0
                Dim pageNum As Integer = Reader.NumberOfPages
                For i As Integer = 1 To pageNum
                    position = objaddon.objDSC.ReadPdfFile(ReadPDF, GetText, i) '"Authorised Signature"
                    x = position(0)
                    y = position(1)
                    page = position(2)
                    llx = x ' left to right  
                    lly = y  'Bottom to Top
                    urx = llx + 150
                    ury = lly + 50
                    If x = 0 Or y = 0 Then
                        ErrFlagCount += 1
                        'objaddon.objapplication.StatusBar.SetText("Reading Text is not found in pdf.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Continue For
                    End If
                    If i > 1 Then
                        Dim stremfile As FileStream = New FileStream(FinalPDFwithDSC, FileMode.Open, FileAccess.Read)
                        Reader = New PdfReader(stremfile)
                        File.Delete(FinalPDFwithDSC)
                    End If
                    Dim stamper As PdfStamper = PdfStamper.CreateSignature(Reader, New FileStream(FinalPDFwithDSC, FileMode.Create, FileAccess.ReadWrite), CChar("\0"), Nothing, True)
                    Dim appearance As PdfSignatureAppearance = stamper.SignatureAppearance


                    If Reason = "Y" Then appearance.Reason = objRs.Fields.Item("U_TxtReason").Value
                    If Location = "Y" Then appearance.Location = objRs.Fields.Item("U_TxtLocation").Value
                    If objRs.Fields.Item("Valid Symbol").Value = "Y" Then appearance.Acro6Layers = False Else appearance.Acro6Layers = True

                    appearance.SetVisibleSignature(New iTextSharp.text.Rectangle(llx, lly, urx, ury), i, Nothing)
                    If objRs.Fields.Item("Valid Symbol").Value = "N" And objRs.Fields.Item("Format").Value = "4" Then
                        Dim wid As Single
                        wid = ColumnText.GetWidth(New Phrase(SignerName))
                        ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_LEFT, New Phrase(SignerName), llx - wid, lly + 15, 0)
                    End If

                    ''appearance.Layer4Text = PdfSignatureAppearance.questionMark

                    If objRs.Fields.Item("Format").Value = "1" Then appearance.Layer2Text = SignerName
                    If objRs.Fields.Item("Format").Value = "3" Then appearance.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.NAME_AND_DESCRIPTION

                    Dim [alias] As String = pfxKeyStore.Aliases.Cast(Of String)().FirstOrDefault(Function(entryAlias) pfxKeyStore.IsKeyEntry(entryAlias))
                    Dim privateKey As ICipherParameters = pfxKeyStore.GetKey([alias]).Key
                    Dim pks As IExternalSignature = New PrivateKeySignature(privateKey, DigestAlgorithms.SHA256)
                    MakeSignature.SignDetached(appearance, pks, New Org.BouncyCastle.X509.X509Certificate() {pfxKeyStore.GetCertificate([alias]).Certificate}, Nothing, Nothing, Nothing, 0, CryptoStandard.CMS)
                    Reader.Close()
                    stamper.Close()
                Next

                If ErrFlagCount = pageNum Then
                    objaddon.objapplication.StatusBar.SetText("""" + GetText + """" + " Text is not found in PDF Document.Please check the document...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    objaddon.objapplication.StatusBar.SetText("Document Signed...Please wait signed document gets opened...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Process.Start(FinalPDFwithDSC)
                End If



                'If pdfs.UpdatedSign(objRs.Fields.Item("U_TxtReason").Value, objaddon.objcompany.CompanyName, objRs.Fields.Item("U_TxtLocation").Value, SignerName, True, ReadPDF) Then
                '    objaddon.objapplication.StatusBar.SetText("Document Signed...Please wait signed document gets opened...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '    Process.Start(FinalPDFwithDSC)
                'Else
                '    objaddon.objapplication.StatusBar.SetText("Document not Signed.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'End If

                'Reader.Close()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Digital_Signature " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Sub Create_Digital_Signature_Without_RPT(ByVal ReadFile As String, ByVal TranName As String)
            Try
                Dim SysPath, Strquery, Foldername, rName, DSCPDFFile As String
                objaddon.objapplication.StatusBar.SetText("Creating DSC Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If objaddon.HANA Then
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select ""AttachPath"" from OADP")
                Else
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select AttachPath from OADP")
                End If
                SysPath = SysPath.Remove(SysPath.Length - 1)
                'Dim directoryFiles = New DirectoryInfo(SysPath)
                ''Dim myFile = (From f In directory.GetFiles() Order By f.LastWriteTime Select f).First()
                'Dim myFile = directoryFiles.GetFiles("*.pdf").OrderByDescending(Function(f) f.LastWriteTime).First()
                'ReadFile = CStr(myFile.FullName)

                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                Else
                    Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                End If
                objRS.DoQuery(Strquery)
                rName = SystemInformation.UserName
                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF\DSC"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                DSCPDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(DSCPDFFile) Then
                    File.Delete(DSCPDFFile)
                End If
                Create_Digital_Signature(objRS.Fields.Item("U_PFXFile").Value, objRS.Fields.Item("U_PFXPass").Value, ReadFile, DSCPDFFile, TranName)
                objRS = Nothing
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Digital_Signature_Without_RPT:" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Function ReadPdfFile(ByVal fileName As String, ByVal searchText As String, Optional PageNum As Integer = 1)
            Dim pages As List(Of Integer) = New List(Of Integer)()
            Dim x, y As Single
            Dim QueryStr As String
            Dim pagecount As Integer, FirstPage As Integer = 1
            Try
                If File.Exists(fileName) Then
                    Dim pdfReader As PdfReader = New PdfReader(fileName)
                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    If objaddon.HANA Then
                        QueryStr = "select ""U_llx"",""U_lly"" from ""@MIPL_ODSC"""
                    Else
                        QueryStr = "select U_llx,U_lly from [@MIPL_ODSC]"
                    End If
                    objRS.DoQuery(QueryStr)
                    'If pdfReader.NumberOfPages <= 2 Then
                    '    FirstPage = pdfReader.NumberOfPages - 1
                    'ElseIf pdfReader.NumberOfPages = 1 Then
                    '    FirstPage = 1
                    'Else
                    '    FirstPage = pdfReader.NumberOfPages - 2
                    'End If

                    Dim strategy As ITextExtractionStrategy = New SimpleTextExtractionStrategy()
                    Dim currentPageText As String = ""

                    If PageNum > 1 Then
                        currentPageText = PdfTextExtractor.GetTextFromPage(pdfReader, PageNum, strategy)
                        If currentPageText.Contains(searchText) Then
                            pagecount = PageNum
                            Dim t = New MyLocationTextExtractionStrategy(searchText, Globalization.CompareOptions.None)
                            Dim ex = PdfTextExtractor.GetTextFromPage(pdfReader, PageNum, t)
                            For Each p In t.myPoints
                                If t.TextToSearchFor = searchText Then
                                    x = p.Rect.Left + CInt(objRS.Fields.Item("U_llx").Value.ToString) '90
                                    y = p.Rect.Bottom + CInt(objRS.Fields.Item("U_lly").Value.ToString) '10
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        For page As Integer = FirstPage To pdfReader.NumberOfPages
                            currentPageText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy)
                            If currentPageText.Contains(searchText) Then
                                pagecount = page
                                Dim t = New MyLocationTextExtractionStrategy(searchText, Globalization.CompareOptions.None)
                                Dim ex = PdfTextExtractor.GetTextFromPage(pdfReader, page, t)
                                For Each p In t.myPoints
                                    If t.TextToSearchFor = searchText Then
                                        x = p.Rect.Left + CInt(objRS.Fields.Item("U_llx").Value.ToString) '90
                                        y = p.Rect.Bottom + CInt(objRS.Fields.Item("U_lly").Value.ToString) '10
                                        Exit For
                                    End If
                                Next
                                If x <> 0 And y <> 0 Then Exit For
                            End If
                        Next
                    End If

                    pdfReader.Close()
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Read_PDF_File_Text:" + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
            Return {x, y, pagecount}
        End Function

        Public Sub Create_Customize_Fields(ByVal objForm As SAPbouiCOM.Form)
            Try
                'objForm = clsModule.objaddon.objapplication.Forms.Item(oFormUID)
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
                'oItem.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                'oItem.Top = objForm.Items.Item("2").Top

                'oItem.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width - 5
                'oItem.Top = objForm.Items.Item("2").Top - objForm.Items.Item("2").Height - 3

                oItem.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
                oItem.Top = objForm.Items.Item("2").Top '- objForm.Items.Item("2").Height - 3

                oItem.Height = objForm.Items.Item("2").Height
                oItem.LinkTo = "2"
                Dim Fieldsize As Size = System.Windows.Forms.TextRenderer.MeasureText("Generate DSC", New Drawing.Font("Arial", 12.0F))
                oItem.Width = Fieldsize.Width - 20
                objForm.Items.Item("bgendsc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("bgendsc").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False)


                oItem = objForm.Items.Add("txtpdf", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                'oItem.Left = objForm.Items.Item("bgendsc").Left - objForm.Items.Item("10000330").Width - 5
                'oItem.Top = objForm.Items.Item("2").Top

                'oItem.Left = objForm.Items.Item("bgendsc").Left + objForm.Items.Item("bgendsc").Width + 5
                'oItem.Top = objForm.Items.Item("2").Top - objForm.Items.Item("2").Height - 5

                oItem.Left = objForm.Items.Item("bgendsc").Left + objForm.Items.Item("bgendsc").Width + 5
                oItem.Top = objForm.Items.Item("2").Top '- objForm.Items.Item("2").Height - 5

                oItem.Width = 100 ' objForm.Items.Item("4").Width - 10
                oItem.Height = 18 ' objForm.Items.Item("4").Height
                oItem.LinkTo = "bgendsc"
                oEdit = oItem.Specific
                Dim TableName As String = objForm.DataSources.DBDataSources.Item(0).TableName
                oEdit.DataBind.SetBound(True, TableName, "U_PDFFile")

            Catch ex As Exception
            End Try
        End Sub

        Public Sub Resize_Customize_Fields(ByVal objForm As SAPbouiCOM.Form)
            Try
                Try
                    oButton = objForm.Items.Item("bgendsc").Specific
                    oEdit = objForm.Items.Item("txtpdf").Specific
                    oButton.Item.Enabled = True
                Catch ex As Exception
                    Exit Sub
                End Try
                If objaddon.HANA Then
                    GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select ""U_RptToPdf"" from ""@MIPL_ODSC"" ") 'where ifnull(""U_RptToPdf"",'N')='N'
                Else
                    GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select U_RptToPdf from [@MIPL_ODSC] ") 'where isnull(U_RptToPdf,'N')='N'
                End If
                If GetEnabledRpttoPDF = "Y" Then
                    'oButton.Item.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                    'oButton.Item.Top = objForm.Items.Item("10000330").Top

                    oButton.Item.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
                    oButton.Item.Top = objForm.Items.Item("2").Top '- objForm.Items.Item("2").Height - 3

                    oEdit.Item.Visible = False
                Else
                    oEdit.Item.Visible = True
                    ''oButton.Item.Left = objForm.Items.Item("10000330").Left - objForm.Items.Item("10000330").Width - 5
                    ''oButton.Item.Top = objForm.Items.Item("10000330").Top
                    'oButton.Item.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 10
                    'oButton.Item.Top = objForm.Items.Item("2").Top - objForm.Items.Item("2").Height - 3
                    ''oEdit.Item.Left = oButton.Item.Left - oEdit.Item.Width - 5 ' 125
                    ''oEdit.Item.Top = objForm.Items.Item("10000330").Top 'Button0.Item.Top

                    oButton.Item.Left = objForm.Items.Item("2").Left + objForm.Items.Item("2").Width + 5
                    oButton.Item.Top = objForm.Items.Item("2").Top '- objForm.Items.Item("2").Height - 3
                    'oEdit.Item.Left = objForm.Items.Item("bgendsc").Left + objForm.Items.Item("bgendsc").Width + 5
                    'oEdit.Item.Top = objForm.Items.Item("2").Top - objForm.Items.Item("2").Height - 3

                    oEdit.Item.Left = objForm.Items.Item("bgendsc").Left + objForm.Items.Item("bgendsc").Width + 5
                    oEdit.Item.Top = objForm.Items.Item("2").Top '- objForm.Items.Item("2").Height - 3
                End If

            Catch ex As Exception

            End Try
        End Sub

    End Class
End Namespace