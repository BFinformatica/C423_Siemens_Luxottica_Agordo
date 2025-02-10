<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Report.aspx.cs" Inherits="NewBfWeb.Pagine.Report" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
   <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
       <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
       <div class="row mt-4 align-items-center">
           <div class="col-0">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["stazioni"]); %>:</div>
                    </div>
                   <asp:DropDownList CssClass="form-control" AutoPostBack="true" ID="stazioni" DataValueField="Code" DataTextField="Description" runat="server" OnSelectedIndexChanged="stazioni_SelectedIndexChanged"></asp:DropDownList>
                </div>
            </div>
           <div class="col-0 ml-1">
                <div class="input-group ml-1">
                    <div class="input-group-prepend">
                        <div class="input-group-text"><% Response.Write(lang["tipo_report"]); %>:</div>
                    </div>
                    <asp:DropDownList CssClass="form-control" AutoPostBack="true" ID="tipo" runat="server" OnSelectedIndexChanged="tipo_SelectedIndexChanged" DataTextField="Description" DataValueField="Code"></asp:DropDownList>
                </div>
            </div>
            <div class="col-0 ml-1" id="parent_anno" runat="server">
                 <div class="input-group ml-1">
                     <div class="input-group-prepend">
                         <div class="input-group-text" id="title_anno" runat="server"></div>
                     </div>
                     <asp:DropDownList CssClass="form-control" AutoPostBack="true" ID="anno" runat="server" OnSelectedIndexChanged="anno_SelectedIndexChanged"></asp:DropDownList>
                 </div>
             </div>
            <div class="col-0 ml-1" id="parent_mese" runat="server">
                 <div class="input-group ml-1">
                     <div class="input-group-prepend">
                         <div class="input-group-text" id="title_mese" runat="server"></div>
                     </div>
                     <asp:DropDownList CssClass="form-control" AutoPostBack="true" ID="mese" runat="server" OnSelectedIndexChanged="mese_SelectedIndexChanged" DataValueField="Index" DataTextField="Testo"></asp:DropDownList>
                 </div>
             </div>
            <div class="col-0 ml-1" id="parent_inizio" runat="server">
                 <div class="input-group ml-1">
                     <div class="input-group-prepend">
                         <div class="input-group-text" id="title_inizio" runat="server"></div>
                     </div>
                     <asp:TextBox CssClass="form-control" AutoPostBack="true" ID="inizio" runat="server" TextMode="Date" OnTextChanged="inizio_TextChanged"></asp:TextBox>
                 </div>
             </div>
            <div class="col-0 ml-1" id="parent_fine" runat="server">
                 <div class="input-group ml-1">
                     <div class="input-group-prepend">
                         <div class="input-group-text" id="title_fine" runat="server"></div>
                     </div>
                     <asp:TextBox CssClass="form-control" AutoPostBack="true" ID="fine" runat="server" TextMode="Date"></asp:TextBox>
                 </div>
             </div>
            <div class="col-0 ml-1" id="parent_ora" runat="server">
                 <div class="input-group ml-1">
                     <div class="input-group-prepend">
                         <div class="input-group-text" id="title_ora" runat="server"></div>
                     </div>
                     <asp:DropDownList CssClass="form-control" AutoPostBack="true" ID="ora" runat="server" DataTextField="dt_hour"></asp:DropDownList>
                 </div>
             </div>
            <div class="col-0 ml-2">
                <asp:Button CssClass="btn btn-primary" runat="server" ID="salva" OnClick="salva_Click"/>
            </div>
       </div>
       <asp:Label ID="errore" runat="server" Visible="false"/>
        <div id="cont_explorer" class="row">
            <div class="col-0">
                <div id="treeview_report">
                    <asp:TreeView ID="elenco_report" runat="server" ShowLines="true" OnTreeNodeExpanded="elenco_report_TreeNodeExpanded" OnSelectedNodeChanged="elenco_report_SelectedNodeChanged">
                        <NodeStyle CssClass="NodeStyle" />
                        <HoverNodeStyle CssClass="HoverNodeStyle"/>
                        <SelectedNodeStyle CssClass="SelectedNodeStyle"/>
                        <RootNodeStyle ImageUrl="~/Img/folder_open.png" />
                        <ParentNodeStyle ImageUrl="~/Img/folder_open.png" />
                        <LeafNodeStyle ImageUrl="~/Img/folder_close.png" />
                    </asp:TreeView>
                </div>
            </div>
            <div class="col-0 ml-1">
            <div id="pnl_explorer">
                <p id="titolo_explorer" runat="server" />
                <div id="explorer" runat="server" class="list-group"/>
            </div>
            </div>
        </div>
    </main>
</asp:Content>
