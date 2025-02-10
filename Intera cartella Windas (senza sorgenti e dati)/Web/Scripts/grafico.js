$(document).ready(function () {

    //quando il documento è caricato
    $(".headerFiltro").each(function () {
        $(this).change(function () {
            var column = "#MainContent_dataTable ." + $(this)[0].value;
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
            var cell = "#MainContent_dataTable ." + $(this).attr('class').split(" ")[1];
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
function ShowAll() {
    var celle = "#MainContent_dataTable tr";
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
    var collection = $("#MainContent_dataTable")[0].children[0].children;
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