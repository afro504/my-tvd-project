from django.contrib import admin

# Register your models here.
# Register your models here.
from mytvddata.models import Country, Component, Subcomponent,Indicator,RepositoryIndicator,SurveyProject,SurveyDataset,DocSave,ReportSave,WarehouseScript,ApiData,StoreAPI,LocationCountry
from import_export.admin import ImportExportModelAdmin
 
# Register your models here.
admin.site.register(Country,ImportExportModelAdmin)
admin.site.register(Component, ImportExportModelAdmin)
admin.site.register(Subcomponent, ImportExportModelAdmin)
admin.site.register(Indicator, ImportExportModelAdmin)
admin.site.register(RepositoryIndicator, ImportExportModelAdmin)
admin.site.register(SurveyProject, ImportExportModelAdmin)
admin.site.register(SurveyDataset, ImportExportModelAdmin)
admin.site.register(DocSave, ImportExportModelAdmin)
admin.site.register(ReportSave, ImportExportModelAdmin)
admin.site.register(WarehouseScript, ImportExportModelAdmin)
admin.site.register(ApiData, ImportExportModelAdmin)
admin.site.register(StoreAPI, ImportExportModelAdmin)
admin.site.register(LocationCountry, ImportExportModelAdmin)
 
 