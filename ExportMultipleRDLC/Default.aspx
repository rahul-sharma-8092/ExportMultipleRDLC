<%@ Page Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ExportMultipleRDLC._Default" %>

<%@ Register Assembly="Microsoft.ReportViewer.WebForms" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <main>
        <div>
            <asp:Button ID="BtnExportAllReport" OnClick="BtnExportAllReport_Click" Text="Export All Report to Excel" runat="server" />
        </div>
        <div>&nbsp;</div>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" Width="720" Height="600px" ShowReportBody="true" KeepSessionAlive="true" PageCountMode="Actual" ProcessingMode="Local" ShowExportControls="True" ShowPrintButton="True" ShowZoomControl="True" ShowParameterArea="True" ZoomMode="Percent" ZoomPercent="100" Style="margin-left: auto; margin-right: auto; display: block;" />

        <div>&nbsp;</div>
        <div>&nbsp;</div>

        <rsweb:ReportViewer ID="ReportViewer2" runat="server" Width="720" Height="600px" ShowReportBody="true" KeepSessionAlive="true" PageCountMode="Actual" ProcessingMode="Local" ShowExportControls="True" ShowPrintButton="True" ShowZoomControl="True" ShowParameterArea="True" ZoomMode="Percent" ZoomPercent="100" Style="margin-left: auto; margin-right: auto; display: block;" />
        
        <div>&nbsp;</div>
        <div>&nbsp;</div>

        <rsweb:ReportViewer ID="ReportViewer3" runat="server" Width="720" Height="600px" ShowReportBody="true" KeepSessionAlive="true" PageCountMode="Actual" ProcessingMode="Local" ShowExportControls="True" ShowPrintButton="True" ShowZoomControl="True" ShowParameterArea="True" ZoomMode="Percent" ZoomPercent="100" Style="margin-left: auto; margin-right: auto; display: block;" />
        
        <div>&nbsp;</div>
        <div>&nbsp;</div>

        <rsweb:ReportViewer ID="ReportViewer4" runat="server" Width="720" Height="600px" ShowReportBody="true" KeepSessionAlive="true" PageCountMode="Actual" ProcessingMode="Local" ShowExportControls="True" ShowPrintButton="True" ShowZoomControl="True" ShowParameterArea="True" ZoomMode="Percent" ZoomPercent="100" Style="margin-left: auto; margin-right: auto; display: block;" />
    </main>


    <script type="text/javascript">
        function downloadExcelFromBase64(base64, fileName) {
            var sampleArr = base64ToArrayBuffer(base64);

            var blob = new Blob([sampleArr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            var link = document.createElement('a');
            link.href = window.URL.createObjectURL(blob);
            link.download = fileName;

            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link); 
        }

        function base64ToArrayBuffer(base64) {
            var binaryString = window.atob(base64);
            var binaryLen = binaryString.length;
            var bytes = new Uint8Array(binaryLen);
            for (var i = 0; i < binaryLen; i++) {
                var ascii = binaryString.charCodeAt(i);
                bytes[i] = ascii;
            }
            return bytes;
        }
    </script>

</asp:Content>
