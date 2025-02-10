(function () {
    //setTime();
    setAllarmi();
    $("#dropdown01").click(function () {
        var asd = $("#utenti_dropdown");
        $("#utenti_dropdown").toggle();
    });
})();
//Gestione del fuso orario
//--------------------------------------------------------------------//
function stdTimezoneOffset(date) {
    var jan = new Date(date.getFullYear(), 0, 1);
    var jul = new Date(date.getFullYear(), 6, 1);
    return Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
}

function isDstObserved(date) {
    return date.getTimezoneOffset() < stdTimezoneOffset(date);
}
//--------------------------------------------------------------------//
//var tick;
var tickAllarmi;
function setAllarmi() {
    //Verifico de ci sono allarmi
    $.ajax({
        type: "POST",
        //url: "Sinottico_ajax.svc/DoWork",
        url: "CalcolaAllarmi.aspx?cont=1",
        data: '{"tags":""}',
        contentType: "application/json; charset=utf-8",
        processData: false,
        dataType: "json",
        success: OnSuccessCount,
        error: OnErrorCall
    });
    var ut = new Date();
    var elemento = document.getElementById("date_time");
    elemento.innerText = formatDate(ut);
    tickAllarmi = setTimeout("setAllarmi()", 4900);
}
function OnSuccessCount(response) {
    if (response[0] > 0) {
        $("#notificaAllarmi").show();
        $("#notificaAllarmi")[0].innerText = response[0];
    }
    else
        $("#notificaAllarmi").hide();
    if (response[1] > 0) {
        $("#notificaStati").show();
        $("#notificaStati")[0].innerText = response[1];
    }
    else
        $("#notificaStati").hide();
}
function formatDate(date) {
    var hours = date.getHours();
    if (isDstObserved(date)) {
        hours--;
    }
    var minutes = date.getMinutes();
    minutes = minutes < 10 ? '0' + minutes : minutes;
    hours = hours < 10 ? '0' + hours : hours;
    var strTime = hours + ':' + minutes;
    return strTime + " - " + date.getDate() + "/" + (date.getMonth() + 1) + "/" + date.getFullYear();
}
function OnErrorCall(response) {
    console.log(response.status + " " + response.statusText);
}