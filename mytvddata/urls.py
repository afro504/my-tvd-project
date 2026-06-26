from django.urls import path
from . import views


urlpatterns = [
 
    # =========================
    # AUTH & HOME
    # =========================
    path('', views.coverPage, name='coverPage'),
    path('signup/', views.signup, name='signup'),
    path('signin/', views.signin, name='signin'),
    path('signout/', views.signout, name='signout'),
 
    # =========================
    # USER / COUNTRY
    # =========================
    path('json_user/', views.json_user, name='json_user'),
    path('add_country/', views.add_country, name='add_country'),
    path('countries/', views.countries, name='countries'),
    path('json_country/', views.json_country, name='json_country'),
    path('export_country/', views.export_country, name='export_country'),
    path("preview-json/", views.preview_json, name="preview_json"),
    path("import-json/", views.import_json, name="import_json"),
 
    path('json_location/', views.json_location, name='json_location'),
 
    # =========================
    # COMPONENT / SUBCOMPONENT
    # =========================
    path('component_add/', views.component_add, name='component_add'),
    path('json_component/', views.json_component, name='json_component'),
    path('export_component/', views.export_component, name='export_component'),
 
    path('subcomponent_add/', views.subcomponent_add, name='subcomponent_add'),
    path('export_subcomponent/', views.export_subcomponent, name='export_subcomponent'),
    path('json_subcomponent/', views.json_subcomponent, name='json_subcomponent'),
 
    # =========================
    # INDICATORS
    # =========================
    path('list_indicator/', views.list_indicator, name='list_indicator'),
    path('create_indicator/', views.create_indicator, name='create_indicator'),
 
    # ✅ CORRECTION IMPORTANTE : retirer &lt; &gt;
    path('update_indicator/<int:pk>/', views.update_indicator, name='update_indicator'),
    path('delete_indicator/<int:pk>/', views.delete_indicator, name='delete_indicator'),
 
    path('export_indicator/', views.export_indicator, name='export_indicator'),
    path('json_indicator/', views.json_indicator, name='json_indicator'),
 
    # =========================
    # WAREHOUSE SCRIPT
    # =========================
    path('t_scripts_view/', views.t_scripts_view, name='t_scripts_view'),
    path('add_t_scripts/new/', views.add_new_t_scripts, name='add_new_t_scripts'),
    path('edit_t_scripts/<int:pk>/', views.edit_t_scripts, name='edit_t_scripts'),
    path('delete_t_scripts/<int:pk>/', views.delete_t_scripts, name='delete_t_scripts'),
 
    # =========================
    # SURVEY PROJECT
    # =========================
    path('project_survey/', views.survey_add_project, name='project_survey'),
    path('json_surveyproject/', views.json_surveyproject, name='json_surveyproject'),
 
    # Upload / dataset
    path('upload_survey/', views.simple_upload, name='upload_survey'),
    path('index_survey/', views.survey_add_data, name='index_survey'),
    path('survey_report/', views.survey_report, name='survey_report'),
 
    path('single_survey_page/<int:pk>/', views.single_survey_page, name='single_survey_page'),
 
    path('export_project/', views.export_project_survey, name='export_project_survey'),
    path('export/', views.export_to_excel, name='export_to_excel'),
    path('export_survey_page/<int:pk>/', views.export_survey_page, name='export_survey_page'),
 
    path('json_surveydataset/', views.json_surveydataset, name='json_surveydataset'),
 
    # =========================
    # REPOSITORY
    # =========================
    path('upload_repository/', views.upload_repository, name='upload_repository'),
    path('repository_report/', views.repository_report, name='repository_report'),
    path('datasets_repository/', views.repository_add_data, name='datasets_repository'),

      # ✅ Nouvelle route pour le mapping
    #path('map_fields/', views.map_fields, name='map_fields'),
 
    path('single_repository_page/<int:pk>/', views.single_repository_page, name='single_repository_page'),
 
    path('export_repository/', views.export_repository_to_excel, name='export_repository'),
    path('export_repository_page/<int:pk>/', views.export_repository_page, name='export_repository_page'),
 
    path('json_repository/', views.json_repository, name='json_repository'),
 
    path(
        'export_to_excel_repositoring/<int:by_subcomponent>/<str:end_day>/',
        views.export_to_excel_repositoring,
        name='export_to_excel_repositoring'
    ),
 
    # =========================
    # API DATA
    # =========================
    path('display_api_data/', views.display_api_data, name='display_api_data'),
    path('load-api-json/', views.load_api_json, name='load_api_json'),# OK to store API data to MySQL
  
    path('loads_api/', views.loads_api, name='loads_api'),
    path('fetch_api_stored/', views.fetch_api_stored, name='fetch_api_stored'),
 
    path('json_storeAPI/', views.json_storeAPI, name='json_storeAPI'),
 
    # =========================
    # REPORT / DOCUMENTS
    # =========================
    path('report_upload/', views.report_upload, name='report_upload'),
    path('views_report/', views.views_report, name='views_report'),
    path('index_report/', views.index_report, name='index_report'),
 
    path('edit_report/<int:pk>/', views.editReport, name='edit_report'),
 
    path('doc_import/', views.docSave_upload, name='doc_import'),
 
    path(
        'export_to_excel_API/<str:by_indicator>/<str:end_day>/',
        views.export_to_excel_API,
        name='export_to_excel_API'
    ),

    path("select_api_fields/", views.select_api_fields, name="select_api_fields"),   # Page pour choisir les variables à insérer
    path("api_data/", views.load_api_other, name="api_data_other"), # Page pour lancer l’extraction et afficher les données
   
    path("imported-data/", views.view_imported_data, name="view_imported_data"),


    # CRUD for STAFF
    path("staff/", views.staff_list, name="staff_list"),
    path("staff/create/", views.staff_create, name="staff_create"),
    path("staff/<int:pk>/update/", views.staff_update, name="staff_update"),
    path("staff/<int:pk>/delete/", views.staff_delete, name="staff_delete"),
     # Import / Export STAFF
    path("staff/export/xlsx/", views.staff_export_xlsx, name="staff_export_xlsx"),
    path("staff/import/xlsx/", views.staff_import_xlsx, name="staff_import_xlsx"),
    # Telecharger model XLSX
    path("staff/template/xlsx/", views.staff_template_xlsx, name="staff_template_xlsx"),




     # Tableau de bord pays
    path("country/<int:pk>/dashboard/", views.country_dashboard, name="country_dashboard"),

    # Export PDF
    #path("country/<int:pk>/dashboard/pdf/", views.export_country_pdf, name="country_dashboard_pdf"),

    # Export Word
    path("country/<int:pk>/dashboard/word/", views.export_country_word, name="country_dashboard_word"),


]
 



