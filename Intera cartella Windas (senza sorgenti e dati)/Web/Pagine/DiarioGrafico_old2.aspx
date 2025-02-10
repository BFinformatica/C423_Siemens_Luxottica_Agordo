<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="DiarioGrafico.aspx.cs" Inherits="NewBfWeb.Pagine.DiarioGrafico" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="col-md-9 ml-sm-auto col-lg-11">
        <script src="../Scripts/chart/dygraph.js"></script>
        <script src="../Scripts/chart/crosshair.js"></script>
        <script src="../Scripts/chart/interaction.js"></script>
        <script src="../Scripts/chart/moment.js"></script>
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
                        <div class="input-group-text"><% Response.Write(lang["tipo_medie"]); %>:</div>
                    </div>
                    <asp:DropDownList CssClass="form-control" AutoPostBack="true" class="custom-select" name="tipo_medie" ID="tipo_medie" runat="server" OnSelectedIndexChanged="tipo_medie_SelectedIndexChanged"></asp:DropDownList>
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
                    <asp:TextBox CssClass="form-control" ID="fine" AutoPostBack="true" name="fine" runat="server" TextMode="Date" />
                </div>
            </div>
            <div class="col-0 ml-2">
                <button class="btn btn-primary" id="carica" type="button">Carica Misure</button>
            </div>
            <div class="col-0 ml-1">
                <a href="#" class="btn btn-primary" onclick="ResetFiltro()"><% Response.Write(lang["resetta_filtro"]); %></a>
            </div>
        </div>
        <asp:Label ID="errore" runat="server" Visible="false"/>
        <div class="row mt-1" id="loading1">
            <div class="col-lg-1" style="margin: 0 auto;">
                <div class="lds-roller"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>
            </div>
        </div>
        <div class="row mt-1" id="div_misure">
            <div class="col-0 ml-1">
                <ul id="mis_1" class="list-group"></ul>
            </div>
            <div class="col-0 ml-1">
                <ul id="mis_2" class="list-group"></ul>
            </div>
            <div class="col-0 ml-1">
                <ul id="mis_3" class="list-group"></ul>
            </div>
            <div class="col-0 ml-1">
                <ul id="mis_4" class="list-group"></ul>
            </div>
            <div class="col-0 ml-1">
                <ul id="mis_5" class="list-group"></ul>
            </div>
            <div class="col-0 ml-1">
                <ul id="mis_6" class="list-group"></ul>
            </div>
            <div class="col-0 ml-1">
                <ul id="mis_7" class="list-group"></ul>
            </div>
        </div>
        <div class="row mt-1">
            <div class="col-0 ml-1">
                <button class="btn btn-primary" id="genera_grafico" type="button">Genera Grafico</button>
            </div>
            <div class="col-0 ml-1">
                <button class="btn btn-primary" id="reset_grafico" type="button">Reset Grafico</button>
            </div>
        </div>
        <div class="row mt-1" id="loading2">
            <div class="col-lg-1" style="margin: 0 auto;">
                <div class="lds-roller"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>
            </div>
        </div>
        <div class="row mt-1" style="position:relative;">
            <div id="myChart"></div>
            <div id="legend"></div>
        </div>
        <script type="text/javascript">
            var g;
            var measure;
            var titoloY = "Perc. (%)";
            $(document).ready(function () {
                $('#genera_grafico').hide();
                $('#reset_grafico').hide();
                $('#div_misure').hide();
                $('#loading1').hide();
                $('#loading2').hide();
                $("#carica").click(function () {
                    $('#loading1').show();
                    //richiamo la funzione che mi carica le misure
                    $.ajax({
                        type: "POST",
                        //url: "Sinottico_ajax.svc/DoWork",
                        url: "CalcolaMisure.aspx?stazione=1",
                        data: '{"tags":""}',
                        contentType: "application/json; charset=utf-8",
                        processData: false,
                        dataType: "json",
                        success: OnMisureSuccessCall,
                        error: OnErrorCall
                    });
                });
                $('#reset_grafico').click(function () {
                    g.updateOptions({
                      dateWindow: null,
                      valueRange: null
                    });
                });
                $('#genera_grafico').click(function () {
                    //chiamo la pagina ch mi restituiswce i punti del grafico in base alle misure checcate
                    misure = '&misure=';
                    for (var count = 0; count < 8; count++) {
                        $('#mis_' + count).children().each(function (element) {
                            if (this.children[0].children[0].checked)
                                misure += this.children[0].children[0].id + "-";
                        });
                    }
                    misure = misure.substring(0, misure.length - 1);
                    $('#loading2').show();
                    $.ajax({
                        type: "POST",
                        //url: "Sinottico_ajax.svc/DoWork",
                        url: "CalcolaGrafico.aspx?stazione=1" + misure + "&tipo_medie=" + $("#MainContent_tipo_medie").val() + "&tipo_dati=" + $("#MainContent_tipo").val() + "&inizio=" + $('#MainContent_inizio').val() + "&fine=" + $('#MainContent_fine').val() + "&perc=1",
                        data: '{"tags":""}',
                        contentType: "application/json; charset=utf-8",
                        processData: false,
                        dataType: "text",
                        success: OnGeneraSuccessCall,
                        error: OnErrorCall
                    });
                });
                    
            });
            function OnMisureSuccessCall(response) {
                $('#genera_grafico').show();
                $('#reset_grafico').show();
                $('#div_misure').show();
                $('#loading1').hide();
                var count = 1;
                measure = response;
                measure.forEach(function (element) {
                    var li = document.createElement('li');
                    li.classList.add('list-group-item');
                    var label = document.createElement('label');
                    var checkbox = document.createElement('input');
                    checkbox.id = element.Codice.trim();
                    checkbox.type = 'checkbox';
                    checkbox.classList.add('misura');
                    label.classList.add('switch');
                    label.appendChild(checkbox);
                    var span = document.createElement("span");
                    span.classList.add("slider");
                    label.appendChild(span);
                    var testo = document.createElement('label');
                    testo.textContent = element.Descrizione.Value;
                    testo.classList.add('testo_margin');
                    li.appendChild(label);
                    li.appendChild(testo);
                    $('#mis_' + count).append(li);
                    count++
                    if (count == 8)
                        count = 1;
                });
            }
            function OnGeneraSuccessCall(response) {
                $('#loading2').hide();
                g = new Dygraph(
                    document.getElementById("myChart"),
                    response,
                    {
                        interactionModel: {
                            'mousedown': downV3,
                            'mousemove': moveV3,
                            'mouseup': upV3,
                            'click': clickV3,
                            'mousewheel': scrollV3
                        },
                        drawPoints: true,
                        pointSize: 3,
                        xlabel: 'Data',
                        ylabel: titoloY,
                        showRangeSelector: false,
                        labelsDiv: document.getElementById('legend'),
                        highlightSeriesOpts: {
                            strokeWidth: 3,
                            strokeBorderWidth: 1,
                            highlightCircleSize: 6,
                        },
                        pointClickCallback: function(event, p) {
                            //chiamo la pagina ch mi restituiswce i punti del grafico in base alle misure checcate
                            for (var i = 0; i < measure.length; i++) {
                                if (measure[i].Codice == p.name) {
                                    titoloY = measure[i].Descrizione.Value;
                                    break;
                                }
                            }
                            misure = '&misure=' + p.name;
                            $.ajax({
                                type: "POST",
                                //url: "Sinottico_ajax.svc/DoWork",
                                url: "CalcolaGrafico.aspx?stazione=1" + misure + "&tipo_medie=" + $("#MainContent_tipo_medie").val() + "&tipo_dati=" + $("#MainContent_tipo").val() + "inizio=" + $('#MainContent_inizio').val() + "&fine=" + $('#MainContent_fine').val() + "&perc=0",
                                data: '{"tags":""}',
                                contentType: "application/json; charset=utf-8",
                                processData: false,
                                dataType: "text",
                                success: OnGeneraSuccessCall,
                                error: OnErrorCall
                            });
                        }
                        //legendFormatter: legendFormatter
                    }
                );
            }
        </script>
    </main>
</asp:Content>
