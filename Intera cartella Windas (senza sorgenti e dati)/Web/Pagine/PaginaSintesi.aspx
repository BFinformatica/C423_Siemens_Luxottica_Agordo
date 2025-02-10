<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="PaginaSintesi.aspx.cs" Inherits="NewBfWeb.Pagine.PaginaSintesi" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
        <div class="row mt-4 align-items-center">
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["Anno"]); %>:</div>
                    </div>
                    <asp:DropDownList class="form-control" id="anno" runat="server"/>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["Mese"]); %>:</div>
                    </div>
                    <asp:DropDownList class="form-control" id="mese" runat="server"/>
                </div>
            </div>
            <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["data"]); %>:</div>
                    </div>
                    <input class="form-control" id="inizio" name="inizio" type="date" />
                </div>
            </div>
        </div>
        <div class="row mt-4 align-items-center">
            <div class="col-0 ml-1">
                <nav aria-label="breadcrumb">
                    <ol class="breadcrumb">
                        <%
                            if ((this.Request.QueryString["sintesi"] != null) && (this.Request.QueryString["periodo"] == "sintesi"))
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><i class=\"far fa-calendar-alt\"></i> Sintesi</li>");
                            else
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><a href=\"../Pagine/PaginaSintesi?sintesi=gruppo1&periodo=sintesi\"><i class=\"far fa-calendar-alt\"></i> Sintesi</a></li>");
                            if ((this.Request.QueryString["sintesi"] != null) && (this.Request.QueryString["periodo"] == "anno"))
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><i class=\"far fa-calendar-alt\"></i> Anno</li>");
                            else
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><a href=\"../Pagine/PaginaSintesi?sintesi=gruppo1&periodo=anno\"><i class=\"far fa-calendar-alt\"></i> Anno</a></li>");
                            if ((this.Request.QueryString["sintesi"] != null) && (this.Request.QueryString["periodo"] != null) && (this.Request.QueryString["periodo"] == "mese"))
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><i class=\"far fa-calendar-alt\"></i> Mese</li>");
                            else
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><a href=\"../Pagine/PaginaSintesi?sintesi=gruppo1&periodo=mese\"><i class=\"far fa-calendar-alt\"></i> Mese</a></li>");
                            if ((this.Request.QueryString["sintesi"] != null) && (this.Request.QueryString["periodo"] != null) && (this.Request.QueryString["periodo"] == "giorno"))
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><i class=\"far fa-calendar-alt\"></i> Giorno</li>");
                            else
                                Response.Write("<li class=\"breadcrumb-item active\" aria-current=\"page\"><a href=\"../Pagine/PaginaSintesi?sintesi=gruppo1&periodo=giorno\"><i class=\"far fa-calendar-alt\"></i> Giorno</a></li>");
                        %>
                    </ol>
                </nav>
            </div>
        </div>
    </main>
</asp:Content>
