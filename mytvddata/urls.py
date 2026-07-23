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


    path('indicators/', views.list_indicator, name='list_indicator'),
    path('indicators/create/', views.create_indicator, name='create_indicator'),
    path('indicators/<int:pk>/update/', views.update_indicator, name='update_indicator'),
    path('indicators/<int:pk>/delete/', views.delete_indicator, name='delete_indicator'),
    path('indicators/export/', views.export_indicator, name='export_indicator'),
    path('indicators/import/', views.import_indicator, name='import_indicator'),
    path('indicators/import/preview/', views.import_preview, name='import_preview'),
    path('indicators/import/confirm/', views.confirm_import, name='confirm_import'),
    path("download_indicator_template/", views.download_indicator_template, name="download_indicator_template"),




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

    path("download_survey_template/", views.download_survey_template, name="download_survey_template"),
    path("preview_import_survey/", views.preview_import_survey, name="preview_import_survey"),
    path("confirm_import_survey/", views.confirm_import_survey, name="confirm_import_survey"),
    path("export_surveyimport_report/", views.export_surveyimport_report, name="export_surveyimport_report"),
  
    path("export_import_report_pdf/", views.export_import_report_pdf, name="export_import_report_pdf"),
 
    path('survey_report/single_survey_page/<int:pk>/', views.single_survey_page, name='single_survey_page'),
 
    path('export_project/', views.export_project_survey, name='export_project_survey'),
    path('export/', views.export_to_excel, name='export_to_excel'),
    path('survey_report/export_survey_page/<int:pk>/', views.export_survey_page, name='export_survey_page'),
 
    path('json_surveydataset/', views.json_surveydataset, name='json_surveydataset'),
 
    # =========================
    # REPOSITORY
    # =========================
    path('upload_repository/', views.upload_repository, name='upload_repository'),
    path('repository_report/', views.repository_report, name='repository_report'),
    path('datasets_repository/', views.repository_add_data, name='datasets_repository'),
    path("import_excel_repository/", views.import_excel_repository, name="import_excel_repository"),
    path("save_excel_repository/", views.save_excel_repository, name="save_excel_repository"),
    path("export_import_report/", views.export_import_report, name="export_import_report"),


      # ✅ Nouvelle route pour le mapping
    #path('map_fields/', views.map_fields, name='map_fields'),
 
    path('single_repository_page/<int:pk>/', views.single_repository_page, name='single_repository_page'),
 
    path('export_repository/', views.export_repository_to_excel, name='export_repository'),
    path('export_repository_page/<int:pk>/', views.export_repository_page, name='export_repository_page'),
 
    path('json_repository/', views.json_repository, name='json_repository'),
 
    path(
        'export_to_excel_repositoring/<int:by_country>/<str:by_subcomponent>/',
        views.export_to_excel_repositoring,
        name='export_to_excel_repositoring'
    ),



  
    path('download_repository_template/', views.download_repository_template, name='download_repository_template'),
      path('export_repository_to_excel/', views.export_repository_to_excel, name='export_repository_to_excel'),
    # =========================
    # API DATA 
    # =========================
    
    # API GHO

    path("select_api_gho_fields/", views.select_api_gho_fields, name="select_api_gho_fields"),   # Page pour choisir les variables à insérer
    path("api_gho_data/", views.load_api_gho_other, name="load_api_gho_other"), # Page pour lancer l’extraction et afficher les données
   
    path("imported-gho-data/", views.view_imported_data_gho, name="view_imported_data_gho"),

    path("view-mapping/", views.view_api_field_mapping, name="view_api_field_mapping"),

    path('display_api_data/', views.display_api_data, name='display_api_data'),




    # OTHERS

    path('indicators/api/', views.load_api_json, name='load_api_json'),
    path('indicators/api/import/', views.import_api_data, name='import_api_data'),

   
 
    path('loads_api/', views.loads_api, name='loads_api'),
    path('fetch_api_stored/', views.fetch_api_stored, name='fetch_api_stored'),
 
    path('json_storeAPI/', views.json_storeAPI, name='json_storeAPI'),
 
    # =========================
    # REPORT / DOCUMENTS
    # =========================
    path('report_upload/', views.report_upload, name='report_upload'),
    path('views_report', views.views_report, name='views_report'),
    path('index_report/', views.index_report, name='index_report'),
 
    path('edit_report/<int:pk>/', views.edit_report, name='edit_report'),
 
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

# Import workflow
    path("staff/import/upload/", views.staff_import_xlsx, name="staff_import_upload"),
    path("staff/import/preview/", views.staff_import_xlsx, name="staff_import_preview"),  # même vue, étape preview
    path("staff/import/mapping/", views.staff_import_mapping, name="staff_import_mapping"),



     # Tableau de bord pays
     # Liste des pays
    path("country_list_dashboard/", views.country_list_dashboard, name="country_list_dashboard"),

    path("country/<int:pk>/dashboard/", views.country_dashboard, name="country_dashboard"),

    # Export PDF
    #path("country/<int:pk>/dashboard/pdf/", views.export_country_pdf, name="country_dashboard_pdf"),

    # Export Word
    path("country/<int:pk>/dashboard/word/", views.export_country_word, name="country_dashboard_word"),

    path("country/<int:pk>/excel/", views.export_country_dashboard_excel, name="export_country_dashboard_excel"),


    ### SURVEY
    path("upload_datasurvey/", views.upload_datasurvey, name="upload_datasurvey"),
    path("surveys/", views.survey_datalist, name="survey_datalist"),

    # ✅ COUNTRY LOCATION
    path("geocountry_list/", views.geocountry_list, name="geocountry_list"),
    path("geocountry_import_json/", views.geocountry_import_json, name="geocountry_import_json"),
    path("geocountry_preview/", views.geocountry_preview, name="geocountry_preview"),
    path("geocountry_sync_to_db/", views.geocountry_sync_to_db, name="geocountry_sync_to_db"),

    # ✅ CRUD LOCATION COUNTRY
    path("geocountries/create/", views.geocountry_create, name="geocountry_create"),
    path("geocountries/<int:pk>/update/", views.geocountry_update, name="geocountry_update"),
    path("geocountries/<int:pk>/delete/", views.geocountry_delete, name="geocountry_delete"),

    # ✅ EXPORT
    path("geocountries/export-json/", views.export_json_geocountry, name="export_json_geocountry"),
    path("geocountries/export-csv/", views.export_csv_geocountry, name="export_csv_geocountry"),

    path("geocountries/data/", views.geocountry_data, name="geocountry_data"),

    path("factsheet/", views.regional_factsheet_select, name="regional_factsheet_select"),
    path("factsheet/<int:subcomponent_id>/", views.regional_factsheet_results, name="regional_factsheet_results"),
    path("factsheet/<int:subcomponent_id>/export_word/", views.export_regional_factsheet_word, name="export_regional_factsheet_word"),
    path("factsheet/<int:subcomponent_id>/excel/", views.export_regional_factsheet_excel, name="export_regional_factsheet_excel"),


]

 



