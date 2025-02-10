<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Sinottico.aspx.cs" Inherits="NewBfWeb.Pagine.Sinottico" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
     <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <div id="visualizzatore_svg" runat="server" style="height: 900px;">
            <%
                System.IO.StreamReader leggi = new System.IO.StreamReader(Server.MapPath("/") + "/Img/sinottico.svg");
                Response.Write(leggi.ReadToEnd());
                leggi.Close();
            %>
        </div>
        <!--Carico il javascript per gestire il soinottico-->
        <script src="../Scripts/sinottico.js"></script>
    </main>
</asp:Content>

