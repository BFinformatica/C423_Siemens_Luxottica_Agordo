<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="SettingUtente.aspx.cs" Inherits="NewBfWeb.Pagine.SettingUtente" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <div class="row mt-4 ml-1 mr-1">
            <asp:Table ID="tbl_utenti" CssClass="table table-striped table-bordered table-hover table-sm" runat="server">

            </asp:Table>
            <div class="card mt-1">
                <div class="card-body">
                    <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
                    <h5 class="card-title"><% Response.Write(lang["inserisci_nuovo_elemento"]); %></h5>
                    <div class="row">
                        <div class="col">
                            <asp:TextBox ID="username" AutoPostBack="true" runat="server" CssClass="form-control"/>
                        </div>
                        <div class="col">
                             <asp:TextBox ID="password" TextMode="Password" AutoPostBack="true" runat="server" CssClass="form-control" />
                        </div>
                        <div class="col-4">
                            <div class="input-group ml-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text"><% Response.Write(lang["tipologia_utente"]); %>:</div>
                                </div>
                             <asp:DropDownList ID="tipo" runat="server" CssClass="form-control"></asp:DropDownList>
                            </div>
                        </div>
                        <div class="col">
                            <asp:Button CssClass="btn btn-primary" ID="inserisci" runat="server" OnClick="inserisci_Click" />
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </main>
</asp:Content>
