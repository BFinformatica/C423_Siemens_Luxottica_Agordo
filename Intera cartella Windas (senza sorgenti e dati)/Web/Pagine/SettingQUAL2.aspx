<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="SettingQUAL2.aspx.cs" Inherits="NewBfWeb.Pagine.SettingQUAL2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
   <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
       <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
        <div class="row mt-4">
            <div class="col-auto">
                <div class="input-group mb-2">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["stazioni"]); %>:</div>
                    </div>
                    <asp:DropDownList ID="stazioni" runat="server" OnSelectedIndexChanged="stazioni_SelectedIndexChanged" DataValueField="Code" DataTextField="Description" CssClass="form-control"/>
                </div>
            </div>
        </div>
        <div class="row">
            <asp:Table ID="setting_table" CssClass="table table-striped table-bordered table-hover table-sm" runat="server">

            </asp:Table>
        </div>
    </main>
</asp:Content>
