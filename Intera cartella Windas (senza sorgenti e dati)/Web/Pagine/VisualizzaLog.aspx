<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="VisualizzaLog.aspx.cs" Inherits="NewBfWeb.Pagine.VisualizzaLog" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <div class="row mt-4">
            <asp:Table ID="tbl_logs" CssClass="table table-striped table-bordered table-hover table-sm" runat="server">
            </asp:Table>
        </div>
        <div class="row mt-1 ml-1 mr-1" id="caricamento">
            <img src="../Img/loading2.gif" alt="loading" height="30" width="30"/> <span style="font-size:30px;"> Loading...</span>
        </div>
        <div class="row mt-1 ml-1 mr-1">
            <div id="finestra">
                <div id="contenuto_finestra">
                    <div class="modal-header">
                        <h5 class="modal-title">File di log</h5>
                        <button type="button" class="close" id="btn_close">
                            <span aria-hidden="true">&times;</span>
                        </button>
                    </div>
                    <div class="row">
                        <div id="vis_log">
                            <table class="table table-striped table-bordered table-hover table-sm">
                                <thead>
                                    <tr>
                                        <th>Gravita</th>
                                        <th>Data e ora</th>
                                        <th>Messaggio</th>
                                    </tr>
                                </thead>
                                <tbody id="tbl_log">
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
            <asp:Table ID="tbl_parametri" runat="server" CssClass="table table-stripe table-bordered table-hover table-sm"></asp:Table>
        </div>
        <script>
            var link = "";
            $(document).ready(function () {
                $("#caricamento").hide();
                $("#finestra").hide();
                $("#btn_close").click(function () {
                    $("#finestra").hide();
                });
                $(".log").click(function () {
                    $("#caricamento").show();
                    link = this.dataset.link;
                    setTimeout(CaricaLog, 0);
                });
                $("#btn_close").click(function () {
                    $("#finestra").hide();
                });
            });
            function CaricaLog() {
                $.getJSON(link, function (response) {
                    $('#tbl_log').remove("tr");
                    for (var k in response) {
                        var riga = document.createElement('tr');
                        var gravita = document.createElement('td');
                        var dataora = document.createElement('td');
                        var messaggio = document.createElement('td');
                        if (response[k].gravita == 'Info')
                            riga.classList.add('table-primary')
                        if (response[k].gravita == 'Warning')
                            riga.classList.add('table-warning')
                        if (response[k].gravita == 'Error')
                            riga.classList.add('table-danger')
                        gravita.textContent = response[k].gravita;
                        dataora.textContent = response[k].dataora;
                        dataora.style.minWidth = '150px';
                        messaggio.textContent = response[k].mex;
                        riga.appendChild(gravita);
                        riga.appendChild(dataora);
                        riga.appendChild(messaggio);
                        $('#tbl_log').append(riga);
                    }
                    $("#caricamento").hide();
                    $("#finestra").show();
                });
            }
        </script>
    </main>
</asp:Content>
