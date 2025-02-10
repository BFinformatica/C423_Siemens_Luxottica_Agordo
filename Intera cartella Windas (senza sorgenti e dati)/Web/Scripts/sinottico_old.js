(function () {
    ricalcola();
})();
//Quando la funzione ajax va a buon fine
var tick;
function OnSuccessCall(response) {
    try {
        //Resetto gli elementi
        response.forEach(function (element) {
            switch (element.Tipo) {
                case 0:
                    var allarme = document.getElementById(element.Id);
                    allarme.style.display = "none";
                    var allarme_value = document.getElementById(element.Id + "_value");
                    allarme_value.innerHTML = "";
                    //allarme_value.style.fontSize = "12px";
                    break;
                case 2:
                    var figura = document.getElementById(element.Id);
                    figura.style.fill = element.ColoreDefault;
                    break;
            }
        });
        response.forEach(function (element) {
            if (element.Tipo == 0) {

            }
            switch (element.Tipo) {
                //Allarme
                case 0:
                    var allarme = document.getElementById(element.Id);
                    var allarme_value = document.getElementById(element.Id + "_value");
                    if (element.IsRosso) {
                        allarme.style.display = "inline";
                        allarme_value.innerHTML += element.Descrizione;
                        if (!element.Resizable) {
                            allarme_value.class = "";
                            if (element.Descrizione.includes("<hr>")) {
                                allarme_value.classList.add('testo_allarme2');
                            }
                        }
                    }
                    break;
                //Misura
                case 1:
                    var titolo_misura = document.getElementById(element.Id + "_titolo");
                    titolo_misura.textContent = element.Descrizione + " " + element.UnitaMisura;
                    var misura = document.getElementById(element.Id + "_valore");
                    misura.textContent = element.Value;
                    break;
                //Figura
                case 2:
                    var figura = document.getElementById(element.Id);
                    if (element.IsRosso) {
                        figura.style.fill = "#ff0000";
                    }
                    else if (element.IsGiallo) {
                        figura.style.fill = "#ffff00";
                    }
                    else if (element.IsVerde) {
                        figura.style.fill = "#00ff00";
                    }
                    else {
                        //Se non è rosso, giallo o verde, lo coloro con il colore di default
                        figura.style.fill = element.ColoreDefault;
                    }
                    break;
            }
        });
    }
    catch (err) {
        if (response.toString().toUpper().includes("TIMEOUT")) {
            window.location.href = "../Login.aspx";
        }
    }
}
function OnErrorCall(response) {
    console.log(response.status + " " + response.statusText);
}
function stop() {
    clearTimeout(tick);
}
function ricalcola() {
    //chiamata ajax ad una pagina con codice c#
    $.ajax({
        type: "POST",
        //url: "Sinottico_ajax.svc/DoWork",
        url: "CalcolaTag.aspx?wb=1&stazione=1",
        data: '{"tags":""}',
        contentType: "application/json; charset=utf-8",
        processData: false,
        dataType: "json",
        success: OnSuccessCall,
        error: OnErrorCall
    });
    tick = setTimeout("ricalcola()", 4900);
}