﻿Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.X509
Imports System.Collections
Imports Org.BouncyCastle.Pkcs
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
'Imports iTextSharp.text.pdf.PdfPKCS7
Imports System.IO
Imports iTextSharp.text.xml.xmp
Imports SAPbobsCOM
Imports SAPbouiCOM
'Imports System.Security.Cryptography.X509Certificates
Imports iTextSharp.text
Imports System.Windows.Forms
Imports iTextSharp.text.pdf.security
Imports DigitalSignature.ClsPDFText

Namespace DigitalSignature
    Public Class PDFSigner

        Public Class Cert
            Private path As String = ""
            Private password As String = ""
            Private akp As AsymmetricKeyParameter
            Private chain As Org.BouncyCastle.X509.X509Certificate()

            Public ReadOnly Property Chain1 As Org.BouncyCastle.X509.X509Certificate()
                Get
                    Return chain
                End Get
            End Property

            Public ReadOnly Property Akp1 As AsymmetricKeyParameter
                Get
                    Return akp
                End Get
            End Property

            Public ReadOnly Property Path1 As String
                Get
                    Return path
                End Get
            End Property

            Public Property Password1 As String
                Get
                    Return password
                End Get
                Set(ByVal value As String)
                    password = value
                End Set
            End Property

            Public Function processCert()
                Dim pk
                Try
                    Dim [alias] As String = Nothing
                    Dim pk12 As Pkcs12Store
                    pk12 = New Pkcs12Store(New FileStream(Me.path, FileMode.Open, FileAccess.Read), Me.password.ToCharArray())
                    Dim i As IEnumerator = pk12.Aliases.GetEnumerator() 'pk12.aliases()

                    While i.MoveNext()
                        [alias] = (CStr(i.Current))
                        If pk12.IsKeyEntry([alias]) Then Exit While
                    End While

                    'Me.akp = pk12.getKey([alias]).getKey()
                    Dim ce As X509CertificateEntry() = pk12.GetCertificateChain([alias])
                    Me.chain = New Org.BouncyCastle.X509.X509Certificate(ce.Length - 1) {}

                    For k As Integer = 0 To ce.Length - 1
                        chain(k) = ce(k).Certificate() 'ce(k).getCertificate()
                    Next
                    pk = pk12.GetKey([alias]).Key
                Catch ex As Exception

                End Try
                Return pk
            End Function

            Public Sub New()
            End Sub

            Public Sub New(ByVal cpath As String)
                Me.path = cpath
                Me.processCert()
            End Sub

            Public Sub New(ByVal cpath As String, ByVal cpassword As String)
                Me.path = cpath
                Me.password = cpassword
                Me.processCert()
            End Sub
        End Class

        Class MetaData
            Private info2 As Hashtable = New Hashtable()
            Private info As Dictionary(Of String, String) = New Dictionary(Of String, String)

            Public Property Info1 As Dictionary(Of String, String)
                Get
                    Return info
                End Get
                Set(ByVal value As Dictionary(Of String, String))
                    info = value
                End Set
            End Property

            Public Property Author As String
                Get
                    Return CStr(info("Author"))
                End Get
                Set(ByVal value As String)
                    info.Add("Author", value)
                End Set
            End Property

            Public Property Title As String
                Get
                    Return CStr(info("Title"))
                End Get
                Set(ByVal value As String)
                    info.Add("Title", value)
                End Set
            End Property

            Public Property Subject As String
                Get
                    Return CStr(info("Subject"))
                End Get
                Set(ByVal value As String)
                    info.Add("Subject", value)
                End Set
            End Property

            Public Property Keywords As String
                Get
                    Return CStr(info("Keywords"))
                End Get
                Set(ByVal value As String)
                    info.Add("Keywords", value)
                End Set
            End Property

            Public Property Producer As String
                Get
                    Return CStr(info("Producer"))
                End Get
                Set(ByVal value As String)
                    info.Add("Producer", value)
                End Set
            End Property

            Public Property Creator As String
                Get
                    Return CStr(info("Creator"))
                End Get
                Set(ByVal value As String)
                    info.Add("Creator", value)
                End Set
            End Property

            Public Function getMetaData() As Dictionary(Of String, String)
                Return Me.info
            End Function

            Public Function getStreamedMetaData() As Byte()
                Dim os As MemoryStream = New System.IO.MemoryStream()
                Dim xmp As XmpWriter = New XmpWriter(os, Me.info)
                xmp.Close()
                Return os.ToArray()
            End Function
        End Class

        Class PDFSigner
            Private inputPDF As String = ""
            Private outputPDF As String = ""
            Public myCert As Cert
            Private metadata As MetaData

            Public Sub New(ByVal input As String, ByVal output As String)
                Me.inputPDF = input
                Me.outputPDF = output
            End Sub

            Public Sub New(ByVal input As String, ByVal output As String, ByVal cert As Cert)
                Me.inputPDF = input
                Me.outputPDF = output
                Me.myCert = cert
            End Sub

            Public Sub New(ByVal input As String, ByVal output As String, ByVal md As MetaData)
                Me.inputPDF = input
                Me.outputPDF = output
                Me.metadata = md
            End Sub

            Public Sub New(ByVal input As String, ByVal output As String, ByVal cert As Cert, ByVal md As MetaData)
                Me.inputPDF = input
                Me.outputPDF = output
                Me.myCert = cert
                Me.metadata = md
            End Sub

            Public Sub Sign(ByVal SigReason As String, ByVal SigContact As String, ByVal SigLocation As String, ByVal SignerName As String, ByVal visible As Boolean)
                Try
                    Dim Reason, Location, position As String
                    Dim objRS As SAPbobsCOM.Recordset
                    Dim reader As PdfReader = New PdfReader(Me.inputPDF)
                    Dim st As PdfStamper = PdfStamper.CreateSignature(reader, New FileStream(Me.outputPDF, FileMode.Create, FileAccess.Write), CChar("\0"), Nothing, True)
                    'Dim st As PdfStamper = New PdfStamper(reader, New FileStream(Me.outputPDF, FileMode.Create, FileAccess.Write), CChar("\0"), True)

                    st.MoreInfo = Me.metadata.getMetaData()
                    st.XmpMetadata = Me.metadata.getStreamedMetaData()
                    Dim sap As PdfSignatureAppearance = st.SignatureAppearance
                    objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    If objaddon.HANA Then
                        position = "select ""U_llx"",""U_lly"" from ""@MIPL_ODSC"""
                    Else
                        position = "select U_llx,U_lly from [@MIPL_ODSC]"
                    End If
                    objRS.DoQuery(position)
                    Dim llx, lly, urx, ury As Integer
                    llx = CInt(objRS.Fields.Item("U_llx").Value.ToString) '400 ' left to right  
                    lly = CInt(objRS.Fields.Item("U_lly").Value.ToString) '245  'Bottom to Top
                    urx = llx + 150
                    ury = lly + 50

                    If objaddon.HANA Then
                        Reason = objaddon.objglobalmethods.getSingleValue("select ""U_Reason"" from ""@MIPL_ODSC""")
                        Location = objaddon.objglobalmethods.getSingleValue("select ""U_Location"" from ""@MIPL_ODSC""")
                        If Reason = "Y" Then
                            sap.Reason = SigReason
                        End If
                        If Location = "Y" Then
                            sap.Location = SigLocation
                        End If
                    Else
                        Reason = objaddon.objglobalmethods.getSingleValue("select U_Reason from [@MIPL_ODSC]")
                        Location = objaddon.objglobalmethods.getSingleValue("select U_Location from [@MIPL_ODSC]")
                        If Reason = "Y" Then
                            sap.Reason = SigReason
                        End If
                        If Location = "Y" Then
                            sap.Location = SigLocation
                        End If
                    End If
                    'If objaddon.HANA Then
                    '    ValidSymbol = objaddon.objglobalmethods.getSingleValue("select ""U_ValidSym"" from  ""@MIPL_ODSC"" ")
                    'Else
                    '    ValidSymbol = objaddon.objglobalmethods.getSingleValue("select U_ValidSym from [@MIPL_ODSC]")
                    'End If
                    'If ValidSymbol = "Y" Then
                    '    sap.Acro6Layers = False
                    'Else
                    '    sap.Acro6Layers = True
                    'End If
                    Dim wid As Single
                    wid = ColumnText.GetWidth(New Phrase(SignerName))
                    ColumnText.ShowTextAligned(st.GetOverContent(reader.NumberOfPages), Element.RECTANGLE, New Phrase(SignerName), llx - wid, lly + 15, 0)
                    sap.Layer4Text = PdfSignatureAppearance.questionMark
                    'sap.SetCrypto(Me.myCert.Akp1, Me.myCert.Chain1, Nothing, PdfSignatureAppearance.VERISIGN_SIGNED) 'PdfSignatureAppearance.VERISIGN_SIGNED

                    'If visible Then sap.SetVisibleSignature(New iTextSharp.text.Rectangle(100, 100, 250, 150), 1, Nothing)
                    'If visible Then sap.SetVisibleSignature(New iTextSharp.text.Rectangle(400, 400, 550, 450), 1, "End of Document")
                    If visible Then sap.SetVisibleSignature(New iTextSharp.text.Rectangle(llx, lly, urx, ury), reader.NumberOfPages, "Signature1")

                    reader.Close()
                    st.Close()
                    ' st.Dispose()

                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText(ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
            End Sub

            Public Function UpdatedSign_Old(ByVal SigReason As String, ByVal SigContact As String, ByVal SigLocation As String, ByVal SignerName As String, ByVal visible As Boolean, ByVal ReadPDF As String, Optional PageNum As Integer = 1) As Boolean
                Try
                    Dim Reason, Location, GetText, TranName As String
                    Dim llx, lly, urx, ury As Integer
                    Dim position As Single()
                    Dim x, y As Single
                    Dim page As Integer
                    If objaddon.objapplication.Forms.ActiveForm.Type.ToString = "133" Then   'AR Invoice
                        TranName = "SI"
                    ElseIf objaddon.objapplication.Forms.ActiveForm.Type.ToString = "140" Then  'Delivery
                        TranName = "DC"
                    ElseIf objaddon.objapplication.Forms.ActiveForm.Type.ToString = "142" Then  'Purchase Order
                        TranName = "PO"
                    ElseIf objaddon.objapplication.Forms.ActiveForm.Type.ToString = "179" Then  'AR Credit Memo
                        TranName = "SR"
                    Else
                        Exit Function
                    End If
                    If objaddon.HANA Then
                        GetText = objaddon.objglobalmethods.getSingleValue("Select Top 1 ""U_Textinpdf"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'")
                    Else
                        GetText = objaddon.objglobalmethods.getSingleValue("select Top 1 U_Textinpdf from [@MIPL_DSC1] where U_TranName='" & TranName & "'")
                    End If
                    If Trim(GetText) = "" Then
                        objaddon.objapplication.StatusBar.SetText("Text is not defined in DSC Settings...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If

                    If objaddon.HANA Then
                        Reason = objaddon.objglobalmethods.getSingleValue("select ""U_Reason"" from ""@MIPL_ODSC""")
                        Location = objaddon.objglobalmethods.getSingleValue("select ""U_Location"" from ""@MIPL_ODSC""")
                    Else
                        Reason = objaddon.objglobalmethods.getSingleValue("select U_Reason from [@MIPL_ODSC]")
                        Location = objaddon.objglobalmethods.getSingleValue("select U_Location from [@MIPL_ODSC]")
                    End If

                    Dim reader As PdfReader = New PdfReader(Me.inputPDF)
                    Dim stamper As PdfStamper = PdfStamper.CreateSignature(reader, New FileStream(Me.outputPDF, FileMode.Create, FileAccess.Write), CChar("\0"), Nothing, True)
                    Dim PN As Integer = 2
                    PageNum = 2
                    'For PageN As Integer = PN To PN 'reader.NumberOfPages

                    'Next
                    Dim appearance As PdfSignatureAppearance = stamper.SignatureAppearance

                    If Reason = "Y" Then appearance.Reason = SigReason
                    If Location = "Y" Then appearance.Location = SigLocation
                    appearance.Acro6Layers = True
                    appearance.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.DESCRIPTION


                    'appearance.Image = Image.GetInstance(File.ReadAllBytes("E:\Chitra\Dawn Motors\greentick.jpg"))
                    'appearance.ImageScale = 1

                    appearance.SignDate = DateTime.Now.Date
                    position = objaddon.objDSC.ReadPdfFile(ReadPDF, GetText, PageNum) '"Authorised Signature"
                    x = position(0)
                    y = position(1)
                    page = position(2)
                    llx = x ' left to right  
                    lly = y  'Bottom to Top
                    urx = llx + 150
                    ury = lly + 50
                    If x = 0 Or y = 0 Then
                        objaddon.objapplication.StatusBar.SetText("Text is not found in pdf.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                    If visible Then
                        appearance.SetVisibleSignature(New iTextSharp.text.Rectangle(llx, lly, urx, ury), reader.NumberOfPages, Nothing) 'page
                    End If

                    Dim wid As Single
                    wid = ColumnText.GetWidth(New Phrase(SignerName))
                    ColumnText.ShowTextAligned(stamper.GetOverContent(page), Element.RECTANGLE, New Phrase(SignerName), llx - wid, lly + 15, 0)
                    'ColumnText.ShowTextAligned(stamper.GetOverContent(page + 1), Element.RECTANGLE, New Phrase(SignerName), llx - wid, lly + 15, 0)

                    'appearance.Layer4Text = PdfSignatureAppearance.questionMark
                    'appearance.Layer2Text = "" '"Signed On" + "12/09/2023"

                    Dim pk = Me.myCert.processCert
                    Dim es As IExternalSignature = New PrivateKeySignature(pk, DigestAlgorithms.SHA256) '"SHA-1""SHA-256"

                    MakeSignature.SignDetached(appearance, es, myCert.Chain1, Nothing, Nothing, Nothing, 0, CryptoStandard.CADES)

                    stamper.Close()
                    Return True
                Catch ex As Exception
                    objaddon.objapplication.SetStatusBarMessage(ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    Return False
                End Try
            End Function

            Public Function SignatureIncorporate_Test(ByVal pdfFilePath As String, ByVal destinationPath As String, ByVal pfxFilePath As String, ByVal pfxPassword As String)
                Try
                    'Dim pdfFilePath As String = "\\newton.tmicloud.net\DB4SHARE\OEC_TEST\Attachments\chitra\chitra\PDF\TestPDF.pdf" ' "C:\Users\Admin\Documents\MyPDF1.pdf"
                    'Dim destinationPath As String = "\\newton.tmicloud.net\DB4SHARE\OEC_TEST\Attachments\chitra\chitra\PDF\DSC\TestPDF.pdf" ' "C:\Users\Admin\Documents\MyPDF_Signed1.pdf"
                    Dim pdfReader As PdfReader = New PdfReader(pdfFilePath)
                    'Dim pfxFilePath As String = "E:\Chitra\Dawn Motors\CK.pfx"
                    'Dim pfxPassword As String = "Teddy@123"
                    Dim pfxKeyStore As Pkcs12Store = New Pkcs12Store(New FileStream(pfxFilePath, FileMode.Open, FileAccess.Read), pfxPassword.ToCharArray())
                    Dim page As Integer = pdfReader.NumberOfPages
                    Dim Reason, Location, GetText, TranName As String
                    Dim llx, lly, urx, ury As Integer
                    Dim position As Single()
                    Dim x, y As Single
                    'Dim page As Integer

                    For i As Integer = 1 To page
                        If i > 1 Then
                            Dim stremfile As FileStream = New FileStream(destinationPath, FileMode.Open, FileAccess.Read)
                            pdfReader = New PdfReader(stremfile)
                            File.Delete(destinationPath)
                        End If

                        Dim signedPdf As FileStream = New FileStream(destinationPath, FileMode.Create, FileAccess.ReadWrite)
                        Dim pdfStamper As PdfStamper = PdfStamper.CreateSignature(pdfReader, signedPdf, vbNullChar, Nothing, True)
                        Dim signatureAppearance As PdfSignatureAppearance = pdfStamper.SignatureAppearance
                        signatureAppearance.Reason = "Digital Signature Reason"
                        signatureAppearance.Location = "Digital Signature Location"
                        signatureAppearance.Acro6Layers = False
                        'Dim x As Single = 360
                        'Dim y As Single = 130
                        position = objaddon.objDSC.ReadPdfFile(pdfFilePath, "Authorised", i) '"Authorised Signature"
                        x = position(0)
                        y = position(1)
                        page = position(2)
                        llx = x ' left to right  
                        lly = y  'Bottom to Top
                        urx = llx + 150
                        ury = lly + 50
                        If x = 0 Or y = 0 Then
                            objaddon.objapplication.StatusBar.SetText("Text is not found in pdf.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Continue For
                        End If
                        signatureAppearance.SetVisibleSignature(New iTextSharp.text.Rectangle(llx, lly, urx, ury), i, Nothing)
                        Dim wid As Single
                        wid = ColumnText.GetWidth(New Phrase("CK"))
                        ColumnText.ShowTextAligned(pdfStamper.GetOverContent(i), Element.RECTANGLE, New Phrase("CK"), llx - wid, lly + 15, 0)

                        signatureAppearance.Acro6Layers = False
                        signatureAppearance.Layer4Text = PdfSignatureAppearance.questionMark
                        'signatureAppearance.SetVisibleSignature(New iTextSharp.text.Rectangle(x, y, x + 150, y + 50), i, Nothing)
                        Dim [alias] As String = pfxKeyStore.Aliases.Cast(Of String)().FirstOrDefault(Function(entryAlias) pfxKeyStore.IsKeyEntry(entryAlias))
                        Dim privateKey As ICipherParameters = pfxKeyStore.GetKey([alias]).Key
                        Dim pks As IExternalSignature = New PrivateKeySignature(privateKey, DigestAlgorithms.SHA256)
                        MakeSignature.SignDetached(signatureAppearance, pks, New Org.BouncyCastle.X509.X509Certificate() {pfxKeyStore.GetCertificate([alias]).Certificate}, Nothing, Nothing, Nothing, 0, CryptoStandard.CMS)
                        pdfReader.Close()
                        pdfStamper.Close()
                    Next
                    Return True
                    'Console.WriteLine("PDF signed successfully!")
                Catch ex As Exception

                End Try
            End Function

        End Class


    End Class
End Namespace