<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="SettingSinottico.aspx.cs" Inherits="NewBfWeb.Pagine.SettingSinottico" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <main role="main" class="pl-4 col-md-9 ml-sm-auto col-lg-11">
        <% NewBfWeb.classi.Languages lang = new NewBfWeb.classi.Languages(); %>
        <div class="row mt-4 ml-1 mr-1">
            <div class="col-0">
                <button type="button" id="add_new" class="btn btn-sm btn-primary"><i class="fas fa-plus"></i> <% Response.Write(lang["inserisci_nuovo_elemento"]); %></button>
            </div>
        </div>
        <div class="row mt-1 ml-1 mr-1">
            <div id="finestra">
                <div id="contenuto_finestra">
                    <div class="row">
                        <div class="col-0">
                            <h1><% Response.Write(lang["inserisci_nuovo_elemento"]); %>:</h1>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-6">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["placeholder_sinottico_verde"]); %>:</div>
                                </div>
                                <input type="text" name="verde" placeholder="Ex. xxx,yyy,zzz" class="form-control" id="verde"/>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-6">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["placeholder_sinottico_giallo"]); %>:</div>
                                </div>
                                <input type="text" name="giallo" placeholder="Ex. xxx,yyy,zzz" class="form-control" id="giallo"/>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-6">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["placeholder_sinottico_rosso"]); %>:</div>
                                </div>
                                <input type="text" name="rosso" placeholder="Ex. xxx,yyy,zzz" class="form-control" id="rosso"/>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-3">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["placeholder_sinottico_id_controllo_svg"]); %>:</div>
                                </div>
                                <input type="text" name="id" placeholder="xyz" class="form-control" id="id"/>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-2">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["tipo_controllo"]); %>:</div>
                                </div>
                                <asp:DropDownList ID="tipo" runat="server"></asp:DropDownList>
                            </div>
                        </div>
                        <div class="col-3 hideable">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["placeholder_unita_misura"]); %>:</div>
                                </div>
                                <input type="text" name="unita_misura" placeholder="xyz" class="form-control" id="unita_misura"/>
                            </div>
                        </div>
                        <div class="col-3 hideable">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["tag_riferimento"]); %>:</div>
                                </div>
                                <input type="text" name="tag_rif" placeholder="xyz" class="form-control" id="tag_rif"/>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-2">
                            <div class="input-group mb-1">
                                <div class="input-group-prepend">
                                    <div class="input-group-text label_desc"><% Response.Write(lang["colore_default"]); %>:</div>
                                </div>
                                <input type="color" name="colore" class="form-control" value="#9b9b9b" id="colore"/>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="col-4">
                            <div class="input-group mb-1">
                                <label class="switch"><input name="resizable" type="checkbox" id="resizable"><span class="slider"></span></label>&nbsp; <% Response.Write(lang["controllo_automaticamente_ridimensionabile"]); %>
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <button class="btn btn-sm btn-primary" type="button" id="salva"><i class="fas fa-save"></i> <% Response.Write(lang["salva"]); %></button>
                        <a class="btn btn-sm btn-primary ml-1" href="#" id="btn_close"><i class="fas fa-times"></i> <% Response.Write(lang["cancel"]); %></a>
                    </div>
                </div>
            </div>
            <asp:Table ID="tbl_parametri" runat="server" CssClass="table table-stripe table-bordered table-hover table-sm"></asp:Table>
        </div>
    </main>
    <script>
        $(document).ready(function () {
            $("#finestra").hide();
            $(".hideable").hide();
            $("#MainContent_tipo").change(function () {
                var idx = this.selectedIndex;
                if (idx == 1)
                    $(".hideable").show();
                else
                    $(".hideable").hide();
            });
            $("#btn_close").click(function () {
                $("#finestra").hide();
            });
            $("#add_new").click(function () {
                $("#finestra").show();
            });
            $("#salva").click(function () {
                $.ajax({
                    type: "POST",
                    url: "InsElemSinottico.aspx?wb=1&verde=" + $("#verde").val()
                        + "&rosso=" + $("#rosso").val()
                        + "&giallo=" + $("#giallo").val()
                        + "&id=" + $("#id").val()
                        + "&tipo=" + $("#MainContent_tipo")[0].selectedIndex
                        + "&unita_misura=" + $("#unita_misura").val()
                        + "&tag_rif=" + $("#tag_rif").val()
                        + "&colore=" + $("#colore").val().replace("#", "")
                        + "&resizable=" + $("#resizable").prop("checked"),
                    data: '{"tags":""}',
                    contentType: "application/json; charset=utf-8",
                    processData: false,
                    dataType: "json",
                    error: OnErrorCall
                });
                $('#finestra').hide();
            });
        });
    </script>
</asp:Content>
