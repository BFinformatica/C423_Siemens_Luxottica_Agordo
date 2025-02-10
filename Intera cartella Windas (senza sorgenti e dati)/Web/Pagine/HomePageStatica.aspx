<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="HomePageStatica.aspx.cs" Inherits="NewBfWeb.Pagine.HomePageStatica" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <style>
            #slider {
                height: 400px;
                min-width: 1280px;
                position: relative;
            }
            #slider > #image-slider, #slider > #image-slider2 {
                background-repeat: no-repeat;
                background-size: cover;
                background-position: center;
                width: 100%;
                height: 100%;
                position: absolute;
                top: 0;
            }
            #slider-destra, #slider-sinistra {
                position: absolute;
                top: 50%;
                color: #ffffff96;
                border-radius: 35px;
                background-color: #00000096;
                width: 35px;
                text-align: center;
                cursor: pointer;
                font-size: 22px;
                line-height: 35px;
            }
            #slider-sinistra {
                left: 10px;
            }
            #slider-destra {
                right: 10px;
            }
            .titolo{
                font-weight:bold;
            }
            .gruppo{
                margin-top: 45px;
            }
        </style>
        <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
        <div class="row" id="slider">
            <div id="image-slider2"></div>
            <div id="image-slider"></div>
            <%--<span id="slider-sinistra">‹</span>
            <span id="slider-destra">›</span>--%>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <h2 class="titolo">La centrale di cogenerazione Albapower</h2>
                <p>La Centrale di cogenerazione è costituita da due impianti cogenerativi alimentati esclusivamente a gas naturale avviati in tempi successivi: il primo impianto a ciclo combinato cogenerativo (GRUPPO 1) di proprietà AlbaPower è stato avviato nel giugno 2007; il secondo impianto cogenerativo (GRUPPO 2) di proprietà Energhe, società energetica del Gruppo Ferrero e gestito da AlbaPower, ha concluso l’avviamento nel mese di agosto 2010.</p>
                <h4 class="gruppo">GRUPPO1</h4>
                <p>L’impianto GRUPPO 1, con potenza complessiva elettrica di 49,95 MWe e termica di 242 MWt è a ciclo combinato. La turbina a gas e dotata di bruciatori a bassa emissione di NOx del tipo DLE (Dry Low Emissions). Il generatore di vapore a recupero (GVR1) è dotato di post combustione (potenza termica complessiva pari a 78 MWt). La turbina a vapore ha un’estrazione controllata di vapore in media pressione (MP), per fornire energia termica alle utenze del teleriscaldamento e dello stabilimento Ferrero. Sul circuito acqua calda per il teleriscaldamento, prima del punto di consegna alla centrale Egea, è posizionato uno stoccaggio di acqua calda, realizzato con quattro serbatoi da 500 m3 cadauno, pari a 2000 m3, utilizzato per coprire le punte giornaliere. In caso di fermata del gruppo cogenerativo il fabbisogno elettrico e soddisfatto dalla rete nazionale alla quale la Centrale e collegata, mentre quello termico da due caldaie ausiliarie (GVA) per la produzione di vapore surriscaldato a 2,3 MPa (23 bar) a 240 °C utilizzanti bruciatori a bassa emissione di NOx. L’acqua del ciclo termoelettrico è prelevata dall’acquedotto comunale, prefiltrata meccanicamente, trattata in un impianto ad osmosi inversa ed in un impianto con resine a scambio ionico che provvede alla demineralizzazione. Una torre evaporativa, in grado di dissipare una potenza di 50 MWt, provvede a condensare il vapore in uscita dalla turbina a vapore, chiudendo il ciclo termodinamico. L’impianto funziona bruciando esclusivamente gas naturale proveniente da un metanodotto ad alta pressione. L’energia elettrica prodotta dall’impianto GRUPPO 1 e immessa alla tensione di 132 kV nella Rete di Trasmissione Nazionale alla quale l’impianto e lo stabilimento Ferrero sono collegati in configurazione entra-esci. La supervisione e gestione dell’impianto e realizzata nella sala controllo di Centrale, presidiata con continuità.</p>
                <h4 class="gruppo">GRUPPO 2</h4>
                <p>Il GRUPPO 2 è un impianto cogenerativo costituito da una turbina a gas (TG2) e da una caldaia a recupero. La turbina a gas è accoppiata direttamente ad un alternatore e ha una potenza elettrica di 6,3 MW; è dotata di un sistema di abbattimento a secco delle emissioni in atmosfera denominato SoLoNOx. Il GVR2 ha una potenza termica complessiva pari a 10,4 MWt. Il vapore generato ad un solo livello di pressione pari a 1,8 MPa, con una temperatura di circa 220°C è destinato alla rete di distribuzione dello stabilimento Ferrero di Alba. Il GVR2 è dotato di economizzatore per il preriscaldo dell’acqua di alimentazione. L’acqua necessaria per il ciclo termico è prelevata dall’impianto di demineralizzazione presente presso il GRUPPO 1. L’energia elettrica è prodotta a 10,5 kV, viene elevata tramite un trasformatore elevatore alla tensione di 30 kV e inviata al sistema elettrico dello stabilimento Ferrero. La sala controllo della Centrale garantisce la supervisione e la gestione dell’Impianto.</p>
            </div>
        </div>
        <script>
            var images = [
                '../img/slide1m.png',
                '../img/slide2m.png',
                '../img/slide3m.png',
                '../img/slide4m.png'
            ]
            var count = 0;
            $(document).ready(function () {
                $("#image-slider").css('background-image', 'url(' + images[count] + ')');
                setTimeout(ToggleImage, 5000);
                $("#slider-sinistra").click(function () {
                    ToggleImage(false);
                });
                $("#slider-destra").click(function () {
                    ToggleImage(true);
                });
            });
            function ToggleImage(destra = true) {
                if (($("#slider-sinistra").prop('disabled')) || ($("#slider-destra").prop('disabled'))) return;
                $("#slider-sinistra").prop('disabled', true);
                $("#slider-destra").prop('disabled', true);
                if (!destra) {
                    if (count == 0)
                        count = images.length - 1;
                    else
                        count--;
                    $('#image-slider').css("right", "");
                    $('#image-slider').animate({ left: 0, width: "toggle" }, 4000, function () {
                        $("#slider-sinistra").prop('disabled', false);
                        $("#slider-destra").prop('disabled', false);
                        $(this).css('background-image', 'url(' + images[count] + ')');
                        setTimeout(ToggleImage, 5000);
                    }).animate({ opacity: 1, left: 0, width: "toggle" }, 0);
                }
                else {
                    if (count >= images.length - 1)
                        count = 0;
                    else
                        count++;
                    $('#image-slider').css("left", "");
                    $('#image-slider').animate({ right: 0, width: "toggle" }, 4000, function () {
                        $("#slider-sinistra").prop('disabled', false);
                        $("#slider-destra").prop('disabled', false);
                        $(this).css('background-image', 'url(' + images[count] + ')');
                        setTimeout(ToggleImage, 5000);
                    }).animate({ opacity: 1, right: 0, width: "toggle" }, 0);
                }
                $('#image-slider2').css('background-image', 'url(' + images[count] + ')');
            }
</script>
    </main>
</asp:Content>
