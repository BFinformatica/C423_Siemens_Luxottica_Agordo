<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="SettingLanguages.aspx.cs" Inherits="NewBfWeb.Pagine.SettingLanguages" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
   <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <div class="row mt-4">
            <div class="col-10">
                <div class="input-group">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><i class="fas fa-search"></i></div>
                    </div>
                    <input type="text" id="search" class="form-control"/>
                </div>
            </div>
        </div>
        <div class="row mt-1">
           <div class="col-10">
               <asp:Table ID="tabella_lang" runat="server" CssClass="table table-striped table-bordered table-hover table-sm" />
           </div>
           <div id="col-0">
               <table class="table table-striped table-bordered table-hover table-sm">
                   <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
                   <tr>
                       <th><% Response.Write(lang["inserisci_nuova_lingua"]); %></th>
                   </tr>
                   <tr>
                       <td>
                           <asp:TextBox ID="lingua" runat="server" AutoPostBack="true" CssClass="form-control" /></td>
                   </tr>
                   <tr>
                       <td>
                           <asp:Button ID="new_lang" CssClass="btn btn-sm btn-primary" runat="server" OnClick="new_lang_Click" AutoPostBack="true" /></td>
                   </tr>
               </table>
               <p>&nbsp;</p>
               <table class="table table-striped table-bordered table-hover table-sm">
                   <tr>
                       <th><% Response.Write(lang["inserisci_nuova_stringa"]); %></th>
                   </tr>
                   <tr>
                       <td>
                           <asp:TextBox ID="stringa" runat="server" AutoPostBack="true" CssClass="form-control" /></td>
                   </tr>
                   <tr>
                       <td colspan="2">
                           <asp:Button ID="new_string" CssClass="btn btn-sm btn-primary" runat="server" OnClick="new_string_Click" AutoPostBack="true" /></td>
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
                $("#MainContent_tabella_lang tr").each(function () {
                    var trovato = false;
                    this.childNodes.forEach(function (element) {
                        if (element.localName == "td") {
                            var testo = element.innerText.toUpperCase();
                            if (testo.includes(value)) {
                                trovato = true;
                                return;
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
            var celle = "#MainContent_tabella_lang tr";
            $(celle).each(function () {
                $(this).show();
            });
        }
    </script>
</asp:Content>
