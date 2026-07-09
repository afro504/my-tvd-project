/*
 * Author: ONDON Auge Wilson
 * Date: 26 October 2025
 * Description:
 *      Demo file for indicator CRUD with AJAX + DataTables + Toast notifications
 **/
$(function () {

  /* -----------------------------
   * Initialisation DataTable
   * ----------------------------- */
  $('#indicator-table').DataTable({
    pageLength: 10,
    dom: 'Bfrtip',
    buttons: [
      { extend: 'excel', className: 'btn btn-outline-success fw-bold' },
      { extend: 'csv', className: 'btn btn-outline-primary fw-bold' },
      { extend: 'pdf', className: 'btn btn-outline-danger fw-bold' },
      { extend: 'print', className: 'btn btn-outline-secondary fw-bold' }
    ]
  });

  /* -----------------------------
   * Fonction Toast Bootstrap
   * ----------------------------- */
  function showToast(message, type = "success") {
    let bgClass = (type === "success") ? "bg-success" : "bg-danger";
    let toastHTML = `
      <div class="toast align-items-center text-white ${bgClass} border-0" role="alert" aria-live="assertive" aria-atomic="true">
        <div class="d-flex">
          <div class="toast-body fw-bold">
            ${message}
          </div>
          <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
        </div>
      </div>`;
    let toastContainer = $("#toast-container");
    if (!toastContainer.length) {
      $("body").append('<div id="toast-container" class="toast-container position-fixed bottom-0 end-0 p-3"></div>');
      toastContainer = $("#toast-container");
    }
    toastContainer.append(toastHTML);
    let toastEl = toastContainer.find(".toast").last()[0];
    let toast = new bootstrap.Toast(toastEl, { delay: 4000 });
    toast.show();
  }

  /* -----------------------------
   * Fonctions AJAX
   * ----------------------------- */
  var loadForm = function () {
    var btn = $(this);
    $.ajax({
      url: btn.attr("data-url"),
      type: 'get',
      dataType: 'json',
      beforeSend: function () {
        $("#modal-indicator .modal-content").html(
          '<div class="p-5 text-center"><i class="fa fa-spinner fa-spin fa-2x text-primary"></i><p>Chargement…</p></div>'
        );
        $("#modal-indicator").modal("show");
      },
      success: function (data) {
        $("#modal-indicator .modal-content").html(data.html_form);
      },
      error: function () {
        showToast("❌ Erreur lors du chargement du formulaire.", "error");
      }
    });
  };

  var saveForm = function (e) {
    e.preventDefault();
    var form = $(this);
    $.ajax({
      url: form.attr("action"),
      data: form.serialize(),
      type: form.attr("method"),
      dataType: 'json',
      success: function (data) {
        if (data.form_is_valid) {
          $("#indicator-table tbody").html(data.html_indicator_list);
          $("#modal-indicator").modal("hide");
          showToast("✅ Indicateur enregistré avec succès.", "success");
        } else {
          $("#modal-indicator .modal-content").html(data.html_form);
        }
      },
      error: function () {
        showToast("❌ Erreur lors de l’enregistrement.", "error");
      }
    });
  };

  /* -----------------------------
   * Bindings
   * ----------------------------- */
  $(".js-create-indicator").click(loadForm);
  $("#indicator-table").on("click", ".js-update-indicator", loadForm);
  $("#indicator-table").on("click", ".js-delete-indicator", loadForm);
  $("#modal-indicator").on("submit", ".js-indicator-form", saveForm);

});
