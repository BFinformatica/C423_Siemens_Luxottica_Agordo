$(document).ready(function () {
    $("#inizio").val(getToday());
    $("#fine").val(getToday());
    ricalcola();
    $("#filtro_allarme").change(Filtra);
    $("#storico").change(disabilitaCheckbox);
    $("#riconosci_all").click(riconosciAll);
});
function Filtra() {
    ShowAll();
    var value = $("#filtro_allarme").val();
    if (value == "") {
        ShowAll();
        return;
    }
    $("#MainContent_tbl_allarmi tr").each(function () {
        if (this.cells[0].nodeName != "TH") {
            var trovato = false;
            for (var i = 0; i < this.cells.length; i++) {
                var testo = this.cells[i].innerText.toUpperCase();
                if (testo.includes(value.toUpperCase()))
                    trovato = true
            }
            if (!trovato)
                $(this).hide();
        }
    });
}
function riconosciAll() {
    $("#corpo_tbl_allarmi tr").each(function () {
        if ((this.cells[0].nodeName != "TH") && ($(this).is(":visible"))) {
            $.ajax({
                type: "POST",
                url: "CalcolaAllarmi.aspx?stazione=1&riconosci_all=1",
                data: '{"tags":""}',
                contentType: "application/json; charset=utf-8",
                processData: false,
                dataType: "json",
                error: OnErrorCall
            });
        }
    });
}
function getToday() {
    var now = new Date();
    var day = ("0" + now.getDate()).slice(-2);
    var month = ("0" + (now.getMonth() + 1)).slice(-2);
    return now.getFullYear() + "-" + (month) + "-" + (day);
}
function ricalcola() {
    $.ajax({
        type: "POST",
        //url: "Sinottico_ajax.svc/DoWork",
        url: "CalcolaAllarmi.aspx?stazione=1&inizio=" + $("#inizio").val()
            + "&fine=" + $("#fine").val()
            + "&storico=" + $("#storico").is(":checked")
            + "&non_riconosciuti=" + $("#non_riconosciuti").is(":checked")
            + "&stati=" + $("#stati").is(":checked")
            + "&allarmi=" + $("#allarmi").is(":checked")
            + "&attivo=" + $("#attivo").is(":checked"),
        data: '{"tags":""}',
        contentType: "application/json; charset=utf-8",
        processData: false,
        dataType: "json",
        success: OnSuccessCall,
        error: OnErrorCall
    });
    tick = setTimeout("ricalcola()", 4900);
}
tick = setTimeout("ricalcola()", 4900);
function OnSuccessCall(response) {
    var corpo = document.getElementById('corpo_tbl_allarmi');
    while (corpo.rows.length > 0) {
        corpo.deleteRow(0);
    }
    response.forEach(function (element) {
        corpo.appendChild(getRiga(element));
    });
    Filtra();
}
function getRiga(element) {
    var riga = document.createElement("tr");
    if (element.Css != "")
        riga.classList.add(element.Css);
    riga.appendChild(getCella(element.Icon, true));
    riga.appendChild(getCella(element.Tipo));
    riga.appendChild(getCella(element.Descrizione));
    riga.appendChild(getCella(element.DataInizio));
    riga.appendChild(getCella(element.OraInizio));
    riga.appendChild(getCella(element.DataFine));
    riga.appendChild(getCella(element.OraFine));
    if (element.IconaRiconosciuto != "")
        riga.appendChild(getCella(element.IconaRiconosciuto, true));
    else
        riga.appendChild(getCella(""));
    riga.appendChild(getCella(element.DataRiconoscimento));
    riga.appendChild(getCella(element.OraRiconoscimento));
    var cella = document.createElement("td");
    if (!element.Riconosciuto) {
        var riconosci = document.createElement("a");
        //Questa parte non è traducibile con le classi c#, quindi va scommentata e\o tradotta all'occorrenza 
        //riconosci.innerText = "Acknowledge alarm";
        riconosci.innerText = "Riconosci allarme";
        riconosci.classList.add("btn");
        riconosci.classList.add("btn-primary");
        riconosci.classList.add("btn-sm");
        //riconosci.onclick += RiconosciAllarme(riga);
        riconosci.addEventListener("click", function () {
            $.ajax({
                type: "POST",
                url: "CalcolaAllarmi.aspx?stazione=1&riconosci=1&data=" + riga.cells[3].innerText + "&ora=" + riga.cells[4].innerText + "&desc=" + riga.cells[2].innerText,
                data: '{"tags":""}',
                contentType: "application/json; charset=utf-8",
                processData: false,
                dataType: "json",
                error: OnErrorCall
            });
        });
        cella.appendChild(riconosci);
    }
    riga.appendChild(cella);
    return riga;
}
function getCella(testo, isIcona = false) {
    var cella = document.createElement("td");
    if (!isIcona)
        cella.innerText = testo;
    else {
        var icona = document.createElement("i");
        icona.innerHTML = testo;
        cella.appendChild(icona);
    }
    return cella;
}
function OnErrorCall(response) {
    console.log(response.status + " " + response.statusText);
}
function ShowAll() {
    var celle = "#MainContent_tbl_allarmi tr";
    $(celle).each(function () {
        $(this).show();
    });
}
function disabilitaCheckbox() {
    $("#non_riconosciuti").prop("disabled", $("#storico").is(":checked"));
    $("#allarmi").prop("disabled", $("#storico").is(":checked"));
    $("#stati").prop("disabled", $("#storico").is(":checked"));
    $("#attivo").prop("disabled", $("#storico").is(":checked"));
}