<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Misure.aspx.cs" Inherits="NewBfWeb.Pagine.Misure" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <table id="tbl_misure" class="mt-4">
            <thead>
                <tr>
                    <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
                    <th rowspan="2" colspan="2">
                        <table style="width:100%;">
                            <tr>
                                <td rowspan="4"><% Response.Write(lang["legenda"]); %>:</td>
                            </tr>
                            <tr class="cell_arancione">
                                <td><% Response.Write(lang["non_valido"]); %></td>
                            </tr>
                            <tr class="cell_giallo">
                                <td><% Response.Write(lang["attenzione"]); %></td>
                            </tr>
                            <tr class="cell_rosso">
                                <td><% Response.Write(lang["allarme"]); %></td>
                            </tr>
                        </table>
                    </th>
                    <th colspan="2"><% Response.Write(lang["valore_istantaneo"]); %></th>
                    <th colspan="4"><% Response.Write(lang["media_oraria"]); %></th>
                    <th colspan="4"><% Response.Write(lang["media_giornaliera"]); %></th>
                </tr>
                <tr>
                    <th><% Response.Write(lang["tal_quale"]); %></th>
                    <th><% Response.Write(lang["elaborato"]); %></th>
                    <th><% Response.Write(lang["precedente"]); %></th>
                    <th><% Response.Write(lang["costruzione"]); %></th>
                    <th><% Response.Write(lang["previsionale"]); %></th>
                    <th><% Response.Write(lang["limite"]); %></th>
                    <th><% Response.Write(lang["precedente"]); %></th>
                    <th><% Response.Write(lang["costruzione"]); %></th>
                    <th><% Response.Write(lang["previsionale"]); %></th>
                    <th><% Response.Write(lang["limite"]); %></th>
                </tr>
            </thead>
        </table>
        <script src="../Scripts/misure.js"></script>
    </main>
</asp:Content>
