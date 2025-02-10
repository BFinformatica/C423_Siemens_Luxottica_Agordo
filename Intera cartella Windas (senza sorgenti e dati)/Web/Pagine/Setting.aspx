<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Setting.aspx.cs" Inherits="NewBfWeb.Pagine.Setting" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
   <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
        <div class="row mt-4">
            <div class="col-10">
                <div class="input-group">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><i class="fas fa-search"></i></div>
                    </div>
                    <input type="text" id="search" class="form-control"/>
                </div>
            </div>
            <div class="col-0">
                <asp:Button Text="Esporta tabelle sql" ID="esporta" runat="server" OnClick="esporta_Click" CssClass="btn btn-sm btn-primary"/>
            </div>
        </div>
        <div class="row mt-1">
            <div class="col-10">
                <asp:Table ID="tbl_parametri" CssClass="table table-striped table-bordered table-hover table-sm" runat="server"></asp:Table>
            </div>
            <div class="col-0">
                <table class="table table-striped table-bordered table-hover table-sm">
                    <tr><th><% Response.Write(lang["inserisci_nuovo_elemento"]); %></th></tr>
                    <tr>
                        <td><asp:TextBox ID="chiave" runat="server" AutoPostBack="true" CssClass="form-control"/></td>
                    </tr>
                    <tr>
                        <td><asp:TextBox ID="valore" runat="server" AutoPostBack="true" CssClass="form-control"/></td>
                    </tr>
                    <tr>
                        <td><asp:Button ID="crea_new" CssClass="btn btn-sm btn-primary" runat="server" OnClick="crea_new_Click" /></td>
                        <!--OnClick="crea_new_Click" -->
                    </tr>
                    <tr>
                        <td><asp:Label ID="errore" runat="server" Visible="false" /></td>
                    </tr>
                </table>
            </div>
        </div>
    </main>
    <script>
        $(document).ready(function () {
            $("#search").change(function () {
                ShowAll();
                var value = $(this).val().toUpperCase();
                if (value == "") {
                    return;
                }
                $("#MainContent_tbl_parametri tr").each(function () {
                    var trovato = false;
                    this.childNodes.forEach(function (element) {
                        if (element.localName == "td") {
                            var testo = element.innerText.toUpperCase();
                            if (testo.includes(value)) {
                                //$($(this).parent()).hide();
                                trovato = true;
                            }
                        }
                        if (element.localName == "th")
                            trovato = true;
                    });
                    if (!trovato)
                        $(this).hide();
                });
            });
        });
        function ShowAll() {
            var celle = "#MainContent_tbl_parametri tr";
            $(celle).each(function () {
                $(this).show();
            });
        }
    </script>
</asp:Content>
