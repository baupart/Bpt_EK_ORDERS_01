Imports System.IO
Imports System.Xml
Imports System.Xml.Xsl

Public Class Bpt_EK_ORDERS_01

    Public Shared Function Bpt_EK_ORDERS_01(var_PfadFiles As String, var_BestellungRequest As String, var_BestellungLieferantennummer As String, var_BestellungBestellnummer As String, var_DbParameterPfad As String)

        Try

            Dim var_PfadBestellungen As String
            Dim var_PfadBestellungenTag As String
            Dim var_PfadXsl As String

            Dim var_XslDateinamePfadStandard As String
            Dim var_XmlDateinameInput As String
            Dim var_XmlDateinameOutput As String

            Dim var_XmlObjekt As New Xml.XmlDocument

            ' Dateipfade und Namen

            var_PfadBestellungen = var_PfadFiles & "\Bestellungen"
            var_PfadBestellungenTag = var_PfadBestellungen & "\" & fct_zeitstempel_tag()
            var_PfadXsl = var_PfadFiles & "\xsl"

            If IO.Directory.Exists(var_PfadBestellungen) Then
            Else
                IO.Directory.CreateDirectory(var_PfadBestellungen)
            End If

            If IO.Directory.Exists(var_PfadBestellungenTag) Then
            Else
                IO.Directory.CreateDirectory(var_PfadBestellungenTag)
            End If

            'Request abspeichern

            var_XmlDateinameInput = var_PfadBestellungenTag & "\" & fct_zeitstempel() & "_" & var_BestellungBestellnummer & "_" & var_BestellungLieferantennummer & "_ERP1_Bestellung_2_1.xml"
            My.Computer.FileSystem.WriteAllText(var_XmlDateinameInput, var_BestellungRequest, True)

            'openTrans_Baupart erzeugen
            var_XslDateinamePfadStandard = var_PfadXsl & "\EK_ORDERS_01.xsl"
            var_XmlDateinameOutput = var_PfadBestellungenTag & "\" & fct_zeitstempel() & "_" & var_BestellungBestellnummer & "_" & var_BestellungLieferantennummer & "_openTransBaupart.xml"

            ' XSL Transformatione Standard opentrans
            Dim xslt As New XslCompiledTransform()
            xslt.Load(var_XslDateinamePfadStandard)
            xslt.Transform(var_XmlDateinameInput, var_XmlDateinameOutput)

            'Wenn alles ordnungsgemäß transformiert wurde, wird der Dateiname zurückgegeben
            Bpt_EK_ORDERS_01 = var_XmlDateinameOutput

        Catch ex As Exception

            'Im Falle eines Fehlers wird der Fehler zurückgegeben, das Keyword "ERROR" wird in der aufrufenden Prozedur abgefragt
            Bpt_EK_ORDERS_01 = "ERROR" & vbCrLf & vbCrLf & ex.ToString

        End Try

    End Function

    Shared Function fct_zeitstempel()

        fct_zeitstempel = Right("0000" & DateTime.Now.Year.ToString(), 4) & Right("00" & DateTime.Now.Month.ToString(), 2) & Right("00" & DateTime.Now.Day.ToString(), 2) & Right("00" & DateTime.Now.Hour.ToString(), 2) & Right("00" & DateTime.Now.Minute.ToString(), 2) & Right("00" & DateTime.Now.Second.ToString(), 2) & Right("000" & DateTime.Now.Millisecond.ToString(), 3)

    End Function

    Shared Function fct_zeitstempel_tag()

        fct_zeitstempel_tag = Right("0000" & DateTime.Now.Year.ToString(), 4) & Right("00" & DateTime.Now.Month.ToString(), 2) & Right("00" & DateTime.Now.Day.ToString(), 2)

    End Function

End Class
