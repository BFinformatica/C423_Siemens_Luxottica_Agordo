<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SinotticoWebForm.aspx.cs" Inherits="NewBfWeb.Pagine.SinotticoWebForm" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="../Content/Site.css" type="text/css" rel="stylesheet"/>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager runat="server">
            <Scripts>
                <asp:ScriptReference Name="jquery" />
            </Scripts>
        </asp:ScriptManager>
        <div id="contenuto">
            <div id="visualizzatore_svg" runat="server" style="height: 785px;">
                <%
                    ///carico il sinottico salvato come svg e lo stampo nella pagina
                    System.IO.StreamReader leggi = new System.IO.StreamReader(Server.MapPath("/") + "/Img/Sinottico.svg");
                    Response.Write(leggi.ReadToEnd());
                    leggi.Close();
                %>
            </div>
            <!--Carico il javascript per gestire il soinottico-->
            <script src="../Scripts/sinottico.js"></script>
        </div>
    </form>
</body>
</html>
