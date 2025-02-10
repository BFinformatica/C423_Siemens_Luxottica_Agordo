$(document).ready(function () {
    //quando il documento è caricato
    $(".headerFiltro > input").each(function () {
        $(this).change(function () {
            var column = "#MainContent_tbl_diario ." + $(this).parent().children()[1].innerText;
            $(column).toggle();
        });
    });
    $(".textFiltro").each(function () {
        $(this).change(function () {
            var value = $(this).val();
            if (value == "") {
                ShowAll();
                return;
            }
            var cell = "#MainContent_tbl_diario ." + $(this).attr('name').split("_")[1];
            $(cell).each(function () {
                if (($(this).is("td")) && ($(this).attr("class").includes("content_cell"))) {
                    var testo = $(this).text();
                    if (!testo.includes(value)) {
                        $($(this).parent()).hide();
                    }
                }
            });
        });
    });
});
function ResetFiltro() {
    var tabella = document.getElementById('MainContent_lista_colonne');
    for (var i = 0; i < tabella.rows.length; i++) {
        for (var y = 0; y < tabella.rows[i].cells.length; y++) {
            if ((tabella.rows[i].cells[y].children[0].classList[0] == "headerFiltro") && (!tabella.rows[i].cells[y].children[0].checked)) {
                tabella.rows[i].cells[y].children[0].checked = true;
                var column = "#MainContent_tbl_diario ." + tabella.rows[i].cells[y].children[0].value;
                $(column).toggle();
            }
            if (tabella.rows[i].cells[y].children[0].classList[0] == "textFiltro") {
                tabella.rows[i].cells[y].children[0].value = "";
            }
        }
    }
    ShowAll();
}
function ShowAll() {
    var celle = "#MainContent_tbl_diario tr";
    $(celle).each(function () {
        $(this).show();
    });
}
function IndexOf(elemento, collection) {
    for (var elem in collection) {
        if (collection[elem].innerText == elemento) {
            return elem;
        }
    }
    return -1;
}
function NascondiColonna(index, nascondi) {
    var collection = $("#MainContent_tbl_diario")[0].children[0].children;
    for (var elem in collection) {
        try {
            if (!isNaN(elem)) {
                if (!nascondi)
                    collection[elem].children[index].style.visibility = "collapse";
                else
                    collection[elem].children[index].style.visibility = "";
            }
        }
        catch (ex) {
            alert(ex);
        }
    }
}