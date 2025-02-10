<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="DiarioTabella.aspx.cs" Inherits="NewBfWeb.Pagine.DiarioTabella" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <div class="row mt-4 align-items-center">
            <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
            <div class="col-0">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["stazioni"]); %>:</div>
                    </div>
                    <asp:DropDownList CssClass="form-control" AutoPostBack="true" class="custom-select" name="stazioni" ID="stazioni" DataValueField="Code" DataTextField="Description" runat="server"></asp:DropDownList>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["tipo_dati"]); %>:</div>
                    </div>
                    <asp:DropDownList CssClass="form-control" AutoPostBack="true" class="custom-select" name="tipo" ID="tipo" runat="server"></asp:DropDownList>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["data_inizio"]); %>:</div>
                    </div>
                    <asp:TextBox CssClass="form-control" ID="inizio" AutoPostBack="true" name="inizio" runat="server" TextMode="Date" />
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["data_fine"]); %>:</div>
                    </div>
                    <asp:TextBox CssClass="form-control" ID="fine" AutoPostBack="true" name="inizio" runat="server" TextMode="Date" />
                </div>
            </div>
            <div class="col-0 ml-2">
                <asp:Button CssClass="btn btn-primary" runat="server" ID="carica" OnClick="carica_Click"/>
            </div>
            <div class="col-0 ml-1">
                <a href="#" class="btn btn-primary" onclick="ResetFiltro()"><% Response.Write(lang["resetta_filtro"]); %></a>
            </div>
        </div>
        <asp:Label ID="errore" runat="server" Visible="false"/>
        <div id="div_filtro" class="row">
            <div class="col">
                <h2><% Response.Write(lang["filtro"]); %>:</h2>
                <div style="max-height:150px; overflow:auto; ">
                     <asp:Table ID="lista_colonne" runat="server" CssClass="table table-striped table-bordered table-hover table-sm"></asp:Table>
                </div>
            </div>
        </div>
        <div id="dati" class="row">
            <div class="col">
                <h2><% Response.Write(lang["dati"]); %>:</h2>
                <div id="contenitore_tabella">
                    <asp:Table ID="tbl_diario" runat="server" CssClass="table table-striped table-bordered table-hover table-sm">
                    </asp:Table>
                </div>
            </div>
        </div>
        <script src="../Scripts/diario.js"></script>
    </main>
</asp:Content>
