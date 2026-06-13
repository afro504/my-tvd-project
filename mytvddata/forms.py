
from django import forms
from django.core.exceptions import ValidationError
from .models import (
    Country, Component, Subcomponent, Indicator,
    RepositoryIndicator, SurveyProject, SurveyDataset,
    DocSave, ReportSave, WarehouseScript, StoreAPI,ApiFieldConfig
)
 
from crispy_forms.helper import FormHelper
from crispy_forms.layout import Layout, Submit, Row, Column
from crispy_forms.bootstrap import FormActions
import django_filters
 
 
# =========================
# WIDGET DATE (HTML5)
# =========================
class DateInput(forms.DateInput):
    input_type = 'date'
 
 
# =========================
# VALIDATION DATE GENERIQUE
# =========================
class DateRangeMixin:
    def clean(self):
        cleaned_data = super().clean()
        start = cleaned_data.get("start")
        end = cleaned_data.get("end")
 
        if start and end and end <= start:
            raise ValidationError("La date de fin doit être supérieure à la date de début")
 
        return cleaned_data
 
 
# =========================
# FORM DATE SIMPLE
# =========================
class DateForm(DateRangeMixin, forms.Form):
    start = forms.DateField(widget=DateInput())
    end = forms.DateField(widget=DateInput())
 
 
# =========================
# COUNTRY
# =========================
class CountryForm(forms.ModelForm):
    class Meta:
        model = Country
        fields = '__all__'
        widgets = {
            'flag': forms.TextInput(attrs={'class': 'form-control'}),
            'cca3': forms.TextInput(attrs={'class': 'form-control'}),
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'official': forms.TextInput(attrs={'class': 'form-control'}),
            'capital': forms.TextInput(attrs={'class': 'form-control'}),
            'subregion': forms.TextInput(attrs={'class': 'form-control'}),
            'area': forms.TextInput(attrs={'class': 'form-control'}),
            'population': forms.NumberInput(attrs={'class': 'form-control'}),
            'languages': forms.Textarea(attrs={'class': 'form-control'}),
            'maps': forms.Textarea(attrs={'class': 'form-control'}),
            'ref_data': forms.TextInput(attrs={'class': 'form-control'}),
            'data_source': forms.Textarea(attrs={'class': 'form-control'}),
            'country_class': forms.TextInput(attrs={'class': 'form-control'}),
        }
 
 
# =========================
# INDICATOR
# =========================
class IndicatorForm(forms.ModelForm):
 
    class Meta:
        model = Indicator
        fields = '__all__'
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        # ✅ Optimisation queryset
        self.fields['subcomponent'].queryset = Subcomponent.objects.select_related('component')
 
        # ✅ Crispy Form
        self.helper = FormHelper()
        self.helper.form_method = 'post'
 
        self.helper.layout = Layout(
            Row(Column('subcomponent')),
            Row(Column('indicator_name')),
            Row(Column('indicator_description')),
            Row(Column('indicator_code'), Column('indicator_target'), Column('category_indicator')),
            Row(Column('indicator_unit'), Column('indicator_frequency'), Column('forecasting_indicator'), Column('ref_data')),
            Row(Column('indicator_source'), Column('performance_indicator'), Column('type_indicator')),
            FormActions(
                Submit('save', 'Save'),
                Submit('cancel', 'Cancel', css_class='btn-danger')
            )
        )
 
 
 
 
class IndicatorFilter(django_filters.FilterSet):
    class Meta:
        model = Indicator
        fields = ['subcomponent']
 


 
class IndicatorSearchForm(forms.Form):
    subcomponent = forms.ModelChoiceField(
        queryset=Subcomponent.objects.all(),
        required=False,
        label="Subcomponent",
        widget=forms.Select(attrs={'class': 'form-control'})
    )
 
    indicator_name = forms.CharField(
        required=False,
        label="Indicator name",
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Search name...'})
    )
 
    start_date = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'})
    )
 
    end_date = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'})
    )
 
    def clean(self):
        cleaned_data = super().clean()
        start = cleaned_data.get("start_date")
        end = cleaned_data.get("end_date")
 
        if start and end and end < start:
            raise forms.ValidationError("End date must be after start date")
 
        return cleaned_data
 
# =========================
# SELECT COMPONENT (OVERRIDE INIT)
# =========================
class SelectcomponentForm(forms.Form):
    by_component = forms.ChoiceField(label='Select component')
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['by_component'].choices = [
            (c.id, c.component_name) for c in Component.objects.all()
        ]
 
 
class SelectSubcomponentForm(forms.Form):
    by_subcomponent = forms.ChoiceField(label='Select subcomponent')
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['by_subcomponent'].choices = [
            (s.id, s.subcomponent_name) for s in Subcomponent.objects.all()
        ]
 
 
# =========================
# DEPENDANT DROPDOWN FIXED ✅
# =========================
class Loadcomponent(forms.Form):
    component = forms.ModelChoiceField(
        queryset=Component.objects.all(),
        empty_label="---------"
    )
 
    subcomponent = forms.ModelChoiceField(
        queryset=Subcomponent.objects.none(),
        empty_label="---------"
    )
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        if 'component' in self.data:
            try:
                component_id = int(self.data.get('component'))
                self.fields['subcomponent'].queryset = Subcomponent.objects.filter(
                    component_id=component_id
                )
            except (ValueError, TypeError):
                pass
 
 

 # =========================
# DOCUMENT UPLOAD
# =========================
class DocSaveForm(forms.ModelForm):
 
    class Meta:
        model = DocSave
        fields = ['name_doc', 'file_doc']
        widgets = {
            'name_doc': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Document name'
            }),
        }
 
    def clean_file_doc(self):
        file = self.cleaned_data.get('file_doc')
 
        if file:
            # ✅ Limite taille fichier (ex: 5MB)
            if file.size > 5 * 1024 * 1024:
                raise forms.ValidationError("Le fichier ne doit pas dépasser 5MB")
 
        return file
 
 
# =========================
# REPORT FORM (UPLOAD + IMAGE)
# =========================
class ReportSaveForm(forms.ModelForm):
 
    class Meta:
        model = ReportSave
        fields = [
            'title_rep', 'author_rep', 'date_rep',
            'summary_rep', 'file_rep', 'img_cp_rep'
        ]
        widgets = {
            'title_rep': forms.Textarea(attrs={
                'class': 'form-control',
                'rows': 2,
                'placeholder': 'Report title'
            }),
            'author_rep': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Author'
            }),
            'date_rep': DateInput(),
            'summary_rep': forms.Textarea(attrs={
                'class': 'form-control',
                'rows': 3,
                'placeholder': 'Summary'
            }),
        }
 
    def clean(self):
        cleaned_data = super().clean()
 
        file_rep = cleaned_data.get('file_rep')
        image = cleaned_data.get('img_cp_rep')
 
        # ✅ Validation taille fichier
        if file_rep and file_rep.size > 10 * 1024 * 1024:
            raise forms.ValidationError("Le fichier PDF ne doit pas dépasser 10MB")
 
        # ✅ Validation image
        if image and image.size > 5 * 1024 * 1024:
            raise forms.ValidationError("L'image ne doit pas dépasser 5MB")
 
        return cleaned_data
 

# =========================
# SURVEY PROJECT
# =========================
class SurveyProjectForm(forms.ModelForm):
 
    class Meta:
        model = SurveyProject
        fields = '__all__'
        widgets = {
            'start_date': DateInput(),
            'end_date': DateInput(),
        }
 
    def clean(self):
        cleaned_data = super().clean()
        start = cleaned_data.get("start_date")
        end = cleaned_data.get("end_date")
 
        if start and end and end <= start:
            raise ValidationError("End date must be greater than start date")
 
        return cleaned_data
 

class SurveyDatasetForm(forms.ModelForm):
 
    class Meta:
        model = SurveyDataset
        fields = [
            'surveyProject', 'quest_code', 'question',
            'response_text', 'response_num',
            'level_1', 'level_2'
        ]
 
        widgets = {
            'surveyProject': forms.Select(attrs={'class': 'form-control'}),
            'quest_code': forms.TextInput(attrs={'class': 'form-control'}),
            'question': forms.TextInput(attrs={'class': 'form-control'}),
            'response_text': forms.TextInput(attrs={'class': 'form-control'}),
            'response_num': forms.NumberInput(attrs={'class': 'form-control'}),
            'level_1': forms.TextInput(attrs={'class': 'form-control'}),
            'level_2': forms.TextInput(attrs={'class': 'form-control'}),
        }
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        # ✅ optimisation légère DB
        self.fields['surveyProject'].queryset = SurveyProject.objects.only('id', 'title_surv')
 
    def clean(self):
        cleaned_data = super().clean()
        text = cleaned_data.get('response_text')
        num = cleaned_data.get('response_num')
 
        # ✅ règle métier forte
        if not text and num is None:
            raise forms.ValidationError("Provide text or numeric response")
 
        return cleaned_data
 

class SelectSurveyForm(forms.Form):
 
    by_survey = forms.ChoiceField(label='Select survey title')
 
    end = forms.DateField(
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        label='Survey end date'
    )
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        # ✅ dynamique + performant
        self.fields['by_survey'].choices = [
            (s.id, s.title_surv)
            for s in SurveyProject.objects.only('id', 'title_surv')
        ]
 
class SurveySearchForm(forms.Form):
 
    survey = forms.ModelChoiceField(
        queryset=SurveyProject.objects.all(),
        required=False,
        widget=forms.Select(attrs={'class': 'form-control'})
    )
 
    question = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Search question'})
    )
 
    level_1 = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control'})
    )
 
    start = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'})
    )
 
    end = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'})
    )
 
    def clean(self):
        cleaned_data = super().clean()
        start = cleaned_data.get('start')
        end = cleaned_data.get('end')
 
        if start and end and end < start:
            raise forms.ValidationError("Invalid date range")
 
        return cleaned_data
 

class SurveyUploadForm(forms.Form):
    file = forms.FileField()
 
    def clean_file(self):
        file = self.cleaned_data.get('file')
 
        if not file.name.endswith('.xlsx'):
            raise forms.ValidationError("Only Excel files allowed")
 
        return file
 
# =========================
# REPOSITORY INDICATOR
# =========================
class RepositoryIndicatorForm(forms.ModelForm):
 
    class Meta:
        model = RepositoryIndicator
        fields = '__all__'
        widgets = {
            'publish_date': DateInput()
        }
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        # ✅ Optimisation
        self.fields['indicator'].queryset = Indicator.objects.select_related('subcomponent')
 
        self.helper = FormHelper()
        self.helper.form_method = 'post'
 
 
class SelectRepositoryForm(forms.Form):
 
    subcomponent = forms.ModelChoiceField(
        queryset=Subcomponent.objects.all(),
        required=False
    )
 
    indicator = forms.ModelChoiceField(
        queryset=Indicator.objects.all(),
        required=False
    )
 
    country = forms.ModelChoiceField(
        queryset=Country.objects.all(),
        required=False
    )
 
    start = forms.DateField(required=False, widget=forms.DateInput(attrs={'type': 'date'}))
    end = forms.DateField(required=False, widget=forms.DateInput(attrs={'type': 'date'}))
 
    def clean(self):
        cleaned_data = super().clean()
 
        if cleaned_data.get("start") and cleaned_data.get("end"):
            if cleaned_data["end"] < cleaned_data["start"]:
                raise forms.ValidationError("Invalid date range")
 
        return cleaned_data
 

# =========================
# DATASET UPLOAD
# =========================
class RepositoryUploadForm(forms.Form):
    excel_file = forms.FileField()
 

 
# =========================
# WAREHOUSE SCRIPT
# =========================
class WarehouseScriptForm(forms.ModelForm):
    class Meta:
        model = WarehouseScript
        fields = '__all__'
 
class S_table_nameForm(forms.Form):
 
    table = forms.ModelChoiceField(
        queryset=WarehouseScript.objects.only('id', 'ws_title'),
        label="Survey",
    )
 
    start = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date'})
    )
 
    end = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date'})
    )
 
    def clean(self):
        cleaned_data = super().clean()
 
        start = cleaned_data.get("start")
        end = cleaned_data.get("end")
 
        if start and end and end < start:
            raise forms.ValidationError("Invalid date range")
 
        return cleaned_data
    

# =========================
# API SELECT (DYNAMIC)
# =========================
class SelectAPIForm(DateRangeMixin, forms.Form):
    by_indicator = forms.ChoiceField(label='Indicator')
 
    start = forms.DateField(widget=DateInput())
    end = forms.DateField(widget=DateInput())
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['by_indicator'].choices = [
            (i.indicator_code, i.indicator_code)
            for i in StoreAPI.objects.all()
        ]
 

class SelectUrlForm(forms.Form):
 
    indicator_code = forms.ChoiceField(label="Indicator code")
 
    start = forms.DateField(required=False, widget=forms.DateInput(attrs={'type': 'date'}))
    end = forms.DateField(required=False, widget=forms.DateInput(attrs={'type': 'date'}))
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        # ✅ dynamique
        self.fields['indicator_code'].choices = [
            (i.indicator_code, i.indicator_code)
            for i in Indicator.objects.only('indicator_code')
        ]
 
    def clean(self):
        cleaned_data = super().clean()
        if cleaned_data.get('start') and cleaned_data.get('end'):
            if cleaned_data['end'] < cleaned_data['start']:
                raise forms.ValidationError("Invalid date range")
        return cleaned_data
 

# =========================
# STORE API
# =========================
class StoreAPIForm(forms.ModelForm):
 
    class Meta:
        model = StoreAPI
        fields = '__all__'
        widgets = {
            'publish_date': DateInput()
        }
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        self.helper = FormHelper()
        self.helper.form_method = 'post'
 
 

# =========================
# COMPONENT FORM
# =========================
class ComponentsForm(forms.ModelForm):
    class Meta:
        model = Component
        fields = ['component_name']
        widgets = {
            'component_name': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Please enter the component name'
            })
        }
 
 
# =========================
# SUBCOMPONENT FORM
# =========================
class SubcomponentsForm(forms.ModelForm):
    class Meta:
        model = Subcomponent
        fields = ['component', 'subcomponent_name']
        widgets = {
            'component': forms.Select(attrs={
                'class': 'form-control'
            }),
            'subcomponent_name': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Please enter the Disease name'
            })
        }
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        # ✅ optimisation des données
        self.fields['component'].queryset = Component.objects.only('id', 'component_name')
 
# =========================
# FORM DYNAMIQUE MODERNE ✅
# =========================
class ComponentForm(forms.Form):
    component = forms.ModelChoiceField(queryset=Component.objects.all())
    subcomponent = forms.ModelChoiceField(queryset=Subcomponent.objects.none())
 
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
 
        if 'component' in self.data:
            try:
                component_id = int(self.data.get('component'))
                self.fields['subcomponent'].queryset = Subcomponent.objects.filter(
                    component_id=component_id
                )
            except:
                pass
 


class SelectUrlForms(forms.Form):
    
    api_url = forms.URLField(
        label="URL de l’API",
        required=True,
        widget=forms.URLInput(attrs={
            "class": "form-control",
            "placeholder": "Entrer l’URL de l’API"
        })
    )


class ApiFieldSelectionForm(forms.Form):
    api_url = forms.URLField(
        label="URL de l’API",
        required=True,
        widget=forms.TextInput(attrs={"placeholder": "https://ghoapi.azureedge.net/api/MALARIA_INDIG"})
    )
    fields = forms.ModelMultipleChoiceField(
        queryset=ApiFieldConfig.objects.all(),
        widget=forms.CheckboxSelectMultiple,
        required=False,
        label="Choisissez les variables à insérer"
    )

    def clean_api_url(self):
        url = self.cleaned_data["api_url"]
        try:
            response = requests.get(url, timeout=5)
            response.raise_for_status()
            data = response.json()

            if "value" not in data or not isinstance(data["value"], list) or len(data["value"]) == 0:
                raise forms.ValidationError("L’API ne contient pas de données exploitables.")
        except Exception as e:
            raise forms.ValidationError(f"Impossible de valider l’URL API : {e}")
        return url

    