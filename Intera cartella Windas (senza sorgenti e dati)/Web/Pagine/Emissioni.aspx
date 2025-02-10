<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Emissioni.aspx.cs" Inherits="NewBfWeb.Pagine.Emissioni" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
        <div class="row">
            <div class="col-lg-12">
                <img src="../Img/albapower_newbig.png" height="690"/>
            </div>
        </div>
        <div class="row mt-1">
            <div class="col-0" style="margin:0 auto;">
                <nav aria-label="breadcrumb">
                    <ol class="breadcrumb">
                        <li class="breadcrumb-item" aria-current="page"><i class="far fa-file-alt"></i> <a href="PaginaDownload.aspx?file_name=../img/limiti_emissivi.pdf&file_name_pulito=limiti_emissivi.pdf">Limiti emissivi</a></li>
                        <li class="breadcrumb-item" aria-current="page"><i class="far fa-file-alt"></i> <a href="PaginaDownload.aspx?file_name=../img/info_sistema.pdf&file_name_pulito=info_sistema.pdf">Curve di correzione</a></li>
                        <li class="breadcrumb-item" aria-current="page"><i class="far fa-file-alt"></i> <a href="PaginaDownload.aspx?file_name=../img/curve_correzione.pdf&file_name_pulito=curve_correzione.pdf">Informazioni di sistema</a></li>
                    </ol>
                </nav>
            </div>
        </div>
    </main>
</asp:Content>
