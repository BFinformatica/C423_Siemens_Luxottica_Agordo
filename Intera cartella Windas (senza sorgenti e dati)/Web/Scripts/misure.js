(function () {
    ricalcola();
})();
var tick;
function ricalcola() {
    $.ajax({
        type: "POST",
        //url: "Sinottico_ajax.svc/DoWork",
        url: "CalcolaMisure.aspx?stazione=1",
        data: '{"tags":""}',
        contentType: "application/json; charset=utf-8",
        processData: false,
        dataType: "json",
        success: OnSuccessCall,
        error: OnErrorCall
    });
    tick = setTimeout("ricalcola()", 4900);
};
tick = setTimeout("ricalcola()", 4900);
function OnSuccessCall(response) {
    var tbl = document.getElementById('tbl_misure');
    while (tbl.rows.length > 2) {
        tbl.deleteRow(2);
    }
    response.forEach(function (element) {
        var tr = document.createElement("tr");
        var tr1 = document.createElement("tr");
        var descrizione = document.createElement("td");
        var progressbar_cell = document.createElement("td");
        var TalQuale = document.createElement("td");
        var TalQuale_info = document.createElement("td");
        var Elaborato = document.createElement("td");
        var Elaborato_info = document.createElement("td");
        var MediaOrariaPrecedente = document.createElement("td");
        var MediaOrariaPrecedente_info = document.createElement("td");
        var MediaOrariaCostruzione = document.createElement("td");
        var MediaOrariaCostruzione_info = document.createElement("td");
        var MediaOrariaPrevisionale = document.createElement("td");
        var MediaOrariaPrevisionale_info = document.createElement("td");
        var MediaOrariaLimite = document.createElement("td");
        var MediaGiornalieraPrecedente = document.createElement("td");
        var MediaGiornalieraPrecedente_info = document.createElement("td");
        var MediaGiornalieraCostruzione = document.createElement("td");
        var MediaGiornalieraCostruzione_info = document.createElement("td");
        var MediaGiornalieraPrevisionale = document.createElement("td");
        var MediaGiornalieraPrevisionale_info = document.createElement("td");
        var MediaGiornalieraLimite = document.createElement("td");
        try {
            descrizione.innerText = element.Descrizione.Value;
            var divProgressContainer = document.createElement("div");
            divProgressContainer.classList.add('progress');
            var divNormaleProgress = document.createElement("div");
            divNormaleProgress.classList.add('progress-bar');
            divNormaleProgress.classList.add('progress-bar-striped');
            divNormaleProgress.classList.add('bg-success');
            divNormaleProgress.setAttribute('aria-valuenow', element.Min + element.Elaborato.Value);
            divNormaleProgress.setAttribute('aria-valuemin', element.Min);
            divNormaleProgress.setAttribute('role', 'progressbar');
            if ((element.SogliaAllarme > 0) && (element.SogliaAttenzione > 0)) {
                divNormaleProgress.setAttribute('aria-valuemax', element.SogliaAttenzione);
                var divAttenzioneProgress = document.createElement("div");
                divAttenzioneProgress.classList.add('progress-bar');
                divAttenzioneProgress.classList.add('progress-bar-striped');
                divAttenzioneProgress.classList.add('progress-bar-animated');
                divAttenzioneProgress.classList.add('bg-warning');
                divAttenzioneProgress.setAttribute('aria-valuenow', element.Min + element.Elaborato.Value);
                divAttenzioneProgress.setAttribute('aria-valuemin', element.SogliaAttenzione);
                divAttenzioneProgress.setAttribute('aria-valuemax', element.SogliaAllarme);
                divAttenzioneProgress.setAttribute('role', 'progressbar');
                var divAllarmeProgress = document.createElement("div");
                divAllarmeProgress.classList.add('progress-bar');
                divAllarmeProgress.classList.add('progress-bar-striped');
                divAllarmeProgress.classList.add('progress-bar-animated');
                divAllarmeProgress.classList.add('bg-danger');
                divAllarmeProgress.setAttribute('aria-valuenow', element.Min + element.Elaborato.Value);
                divAllarmeProgress.setAttribute('aria-valuemin', element.SogliaAllarme);
                divAllarmeProgress.setAttribute('aria-valuemax', element.Max);
                divAllarmeProgress.setAttribute('role', 'progressbar');
                var valMinReale = element.Elaborato.Value - element.Min;
                if (valMinReale <= element.SogliaAttenzione) {
                    divNormaleProgress.style.width = (valMinReale * 100) / element.SogliaAttenzione + '%';
                    divAttenzioneProgress.style.width = divAttenzioneProgress.style.width = '0%';
                }
                else if ((valMinReale > element.SogliaAttenzione) && (valMinReale <= element.SogliaAllarme)) {
                    divNormaleProgress.style.width = '100%';
                    divAttenzioneProgress.style.width = ((valMinReale - element.SogliaAttenzione) * 100) / element.SogliaAllarme + '%';
                    divAllarmeProgress.style.width = '0%';
                }
                else if ((valMinReale > element.SogliaAttenzione) && (valMinReale > element.SogliaAllarme)) {
                    divNormaleProgress.style.width = '100%';
                    divAttenzioneProgress.style.width = '100%';
                    var asd = ((valMinReale - element.SogliaAllarme) * 100) / element.Max;
                    divAllarmeProgress.style.width = ((valMinReale - element.SogliaAllarme) * 100) / element.Max + '%';
                }
                divProgressContainer.appendChild(divNormaleProgress);
                divProgressContainer.appendChild(divAttenzioneProgress);
                divProgressContainer.appendChild(divAllarmeProgress);
            }
            else {
                divNormaleProgress.setAttribute('aria-valuemax', element.Max);
                divNormaleProgress.style.width = ((element.Elaborato.Value - element.Min) * 100) / element.Max + '%';
                divProgressContainer.appendChild(divNormaleProgress);
            }
            divProgressContainer.style.minWidth ='100px';
            progressbar_cell.appendChild(divProgressContainer);
            TalQuale.innerText = SistemaMisura(element.TalQuale.Value);
            TalQuale.classList.add("cell_" + element.TalQuale.Colore);
            TalQuale_info.innerHTML = this.getCellInfo(element.TalQuale, element.IdMisuraDatabase);
            Elaborato.innerText = SistemaMisura(element.Elaborato.Value);
            Elaborato.classList.add("cell_" + element.Elaborato.Colore);
            Elaborato_info.innerHTML = this.getCellInfo(element.TalQuale, element.IdMisuraDatabase);
            if (!element.IsSemioraria) {
                MediaOrariaPrecedente.innerText = SistemaMisura(element.MediaOrariaPrecedente.Value);
                MediaOrariaPrecedente.classList.add("cell_" + element.MediaOrariaPrecedente.Colore);
                MediaOrariaPrecedente_info.innerHTML = this.getCellInfo(element.MediaOrariaPrecedente, element.IdMisuraDatabase);
                MediaOrariaCostruzione.innerText = SistemaMisura(element.MediaOrariaCostruzione.Value);
                MediaOrariaCostruzione.classList.add("cell_" + element.MediaOrariaCostruzione.Colore);
                MediaOrariaCostruzione_info.innerHTML = this.getCellInfo(element.MediaOrariaCostruzione, element.IdMisuraDatabase);
                MediaOrariaPrevisionale.innerText = SistemaMisura(element.MediaOrariaPrevisionale.Value);
                MediaOrariaPrevisionale.classList.add("cell_" + element.MediaOrariaPrevisionale.Colore);
                MediaOrariaPrevisionale_info.innerHTML = this.getCellInfo('', element.IdMisuraDatabase);
                MediaOrariaLimite.innerText = element.LimiteMediaOraria;// + '/' + element.LimiteMediaOraria2;
                MediaOrariaLimite.classList.add("cell_limite");
                MediaGiornalieraPrecedente.innerText = SistemaMisura(element.MediaGiornalieraPrecedenteOraria.Value);
                MediaGiornalieraPrecedente.classList.add("cell_" + element.MediaGiornalieraPrecedenteOraria.Colore);
                MediaGiornalieraPrecedente_info.innerHTML = this.getCellInfo(element.MediaGiornalieraPrecedenteOraria, element.IdMisuraDatabase);
                MediaGiornalieraCostruzione.innerText = SistemaMisura(element.MediaGiornalieraCostruzioneOraria.Value);
                MediaGiornalieraCostruzione.classList.add("cell_" + element.MediaGiornalieraCostruzioneOraria.Colore);
                MediaGiornalieraCostruzione_info.innerHTML = this.getCellInfo(element.MediaGiornalieraCostruzioneOraria, element.IdMisuraDatabase);
                MediaGiornalieraPrevisionale.innerText = SistemaMisura(element.MediaGiornalieraPrevisionaleOraria.Value);
                MediaGiornalieraPrevisionale.classList.add("cell_" + element.MediaGiornalieraPrevisionaleOraria.Colore);
                MediaGiornalieraPrevisionale_info.innerHTML = this.getCellInfo('', element.IdMisuraDatabase);
                MediaGiornalieraLimite.innerText = element.LimiteMediaGiornalieraOraria;// + '/' + element.LimiteMediaGiornalieraOraria2;
                MediaGiornalieraLimite.classList.add("cell_limite");
            }
            else {
                MediaOrariaPrecedente.innerText = SistemaMisura(element.MediaSemiorariaPrecedente.Value);
                MediaOrariaPrecedente.classList.add("cell_" + element.MediaSemiorariaPrecedente.Colore);
                MediaOrariaPrecedente_info.innerHTML = this.getCellInfo(element.MediaSemiorariaPrecedente, element.IdMisuraDatabase);
                MediaOrariaCostruzione.innerText = SistemaMisura(element.MediaSemiorariaCostruzione.Value);
                MediaOrariaCostruzione.classList.add("cell_" + element.MediaSemiorariaCostruzione.Colore);
                MediaOrariaCostruzione_info.innerHTML = this.getCellInfo(element.MediaSemiorariaCostruzione, element.IdMisuraDatabase);
                MediaOrariaPrevisionale.innerText = SistemaMisura(element.MediaSemiorariaPrevisionale.Value);
                MediaOrariaPrevisionale.classList.add("cell_" + element.MediaSemiorariaPrevisionale.Colore);
                MediaOrariaPrevisionale_info.innerHTML = this.getCellInfo('', element.IdMisuraDatabase);
                MediaOrariaLimite.innerText = element.LimiteMediaSemioraria;// + '/' + element.LimiteMediaSemioraria2;
                MediaOrariaLimite.classList.add("cell_limite");
                MediaGiornalieraPrecedente.innerText = SistemaMisura(element.MediaGiornalieraPrecedenteSemioraria.Value);
                MediaGiornalieraPrecedente.classList.add("cell_" + element.MediaGiornalieraPrecedenteSemioraria.Colore);
                MediaGiornalieraPrecedente_info.innerHTML = this.getCellInfo(element.MediaGiornalieraPrecedenteSemioraria, element.IdMisuraDatabase);
                MediaGiornalieraCostruzione.innerText = SistemaMisura(element.MediaGiornalieraCostruzioneSemioraria.Value);
                MediaGiornalieraCostruzione.classList.add("cell_" + element.MediaGiornalieraCostruzioneSemioraria.Colore);
                MediaGiornalieraCostruzione_info.innerHTML = this.getCellInfo(element.MediaGiornalieraCostruzioneSemioraria, element.IdMisuraDatabase);
                MediaGiornalieraPrevisionale.innerText = SistemaMisura(element.MediaGiornalieraPrevisionaleSemioraria.Value);
                MediaGiornalieraPrevisionale.classList.add("cell_" + element.MediaGiornalieraPrevisionaleSemioraria.Colore);
                MediaGiornalieraPrevisionale_info.innerHTML = this.getCellInfo('', element.IdMisuraDatabase);
                MediaGiornalieraLimite.innerText = element.LimiteMediaGiornalieraSemioraria;
                MediaGiornalieraLimite.classList.add("cell_limite");
            }
            descrizione.rowSpan =
                progressbar_cell.rowSpan =
                MediaOrariaLimite.rowSpan =
                MediaGiornalieraLimite.rowSpan = 2;
            tr.appendChild(descrizione);
            tr.appendChild(progressbar_cell);
            tr.appendChild(TalQuale);
            tr.appendChild(Elaborato);
            tr.appendChild(MediaOrariaPrecedente);
            tr.appendChild(MediaOrariaCostruzione);
            tr.appendChild(MediaOrariaPrevisionale);
            tr.appendChild(MediaOrariaLimite);
            tr.appendChild(MediaGiornalieraPrecedente);
            tr.appendChild(MediaGiornalieraCostruzione);
            tr.appendChild(MediaGiornalieraPrevisionale);
            tr.appendChild(MediaGiornalieraLimite);
            tr1.classList.add('rigaPedice');
            tr1.appendChild(TalQuale_info);
            tr1.appendChild(Elaborato_info);
            tr1.appendChild(MediaOrariaPrecedente_info);
            tr1.appendChild(MediaOrariaCostruzione_info);
            tr1.appendChild(MediaOrariaPrevisionale_info);
            tr1.appendChild(MediaGiornalieraPrecedente_info);
            tr1.appendChild(MediaGiornalieraCostruzione_info);
            tr1.appendChild(MediaGiornalieraPrevisionale_info);
            tbl.appendChild(tr);
            tbl.appendChild(tr1);
        }
        catch (error) {
            console.log(response.status + " " + response.statusText);
        }
    });
}
function SistemaMisura(valore) {
    if (valore == -9999)
        return "---";
    return valore;
}
function OnErrorCall(response) {
    console.log(response.status + " " + response.statusText);
}
function getCellInfo(valore, id = '') {
    var div_id = '';
    if (valore == '')
        return '';
    if (valore.ID != '')
        div_id = '<p class="pedice_ID">' + valore.ID + '</p>';
    var div = '<div style="position:relative;"><p class="pedice_Valido">VAL</p>' + div_id + '</div>';
    if (!valore.Valido)
        div = '<div style="position:relative;"><p class="pedice_Err">ERR</p>' + div_id + '</div>';
    return div;
}