from import_export import resources
from .models import SurveyDataset
 
class SurveyResource(resources.ModelResource):
    class meta:
        model = SurveyDataset
 