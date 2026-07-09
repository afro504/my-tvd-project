# Create your models here.
from django.db import models
from django.urls import reverse


# ============================
# TABLE DES PAYS (Country)
# ============================
class Country(models.Model):
    flag = models.SlugField('Flag', blank=True, null=True)                     # Drapeau (slug)
    cca3 = models.CharField('ISO', max_length=300)                             # Code ISO3
    name = models.CharField('Country name', max_length=300)                    # Nom du pays
    official = models.CharField('Official', max_length=250)                    # Nom officiel
    capital = models.CharField('Capital', max_length=250)                      # Capitale
    subregion = models.CharField('Subregion', max_length=250)                  # Sous-région
    area = models.CharField('Area', max_length=250)                            # Superficie
    population = models.FloatField('Population')                               # Population
    languages = models.CharField('Languages', max_length=250)                  # Langues
    maps = models.URLField('Maps', max_length=500)                             # Lien vers carte
    ref_data = models.CharField('Reference of data', max_length=50)            # Référence des données
    data_source = models.URLField('Link', max_length=500)                      # Source des données
    country_class = models.CharField('Country class', max_length=300)          # Classification du pays
    
    def get_absolute_url(self):
        return reverse('mytvddata:country_detail', kwargs={'pk': self.pk})           # URL absolue pour détail

    def __str__(self):
        return self.name

    class Meta:
        db_table = 'country'
        ordering = ['name']                                                    # Tri par nom




# ============================
# TABLE DES PAYS AVEC COORDONNÉES
# ============================
class LocationCountry(models.Model):
    iso3 = models.CharField('ISO', max_length=300)  # Code ISO3 du pays
    name = models.CharField(max_length=100)         # Nom du pays
    latitude = models.FloatField()                  # Latitude géographique
    longitude = models.FloatField()                 # Longitude géographique
    
    class Meta:
        db_table = 'location_country'
        ordering = ['name']  # Tri par nom

    def __str__(self):
        return self.name


# ============================
# TABLE DES VILLES
# ============================
class City(models.Model):
    country = models.CharField('Country name', max_length=300)  # Nom du pays
    iso3 = models.CharField('ISO', max_length=300)              # Code ISO3 du pays
    flag = models.SlugField('Flag', blank=True, null=True)      # Drapeau (slug)
    city_country = models.CharField('City of country', max_length=250)  # Nom de la ville
    city_ascii = models.CharField('City ascii', max_length=250)         # Nom en ASCII
    admin_name = models.CharField('Subregion', max_length=250)          # Région administrative
    latitude = models.FloatField()                                     # Latitude
    longitude = models.FloatField()                                    # Longitude
    capital = models.CharField('Capital', max_length=250, null=True, blank=True)  # Capitale
    population = models.FloatField('Population', null=True, blank=True)           # Population
    
    def __str__(self):
        return self.country

    class Meta:
        db_table = 'list_city'
        ordering = ['country']


# ============================
# TABLE DES COMPOSANTS
# ============================
class Component(models.Model):
    component_name = models.CharField('Component', max_length=50)  # Nom du composant
    
    class Meta:
        db_table = 'list_component'
        ordering = ['component_name']

    def __str__(self):
        return self.component_name


# ============================
# TABLE DES SOUS-COMPOSANTS
# ============================
class Subcomponent(models.Model):
    subcomponent_name = models.CharField('Disease', max_length=250, null=True)  # Nom du sous-composant
    component = models.ForeignKey(Component, on_delete=models.SET_NULL, null=True, related_name='comp_subcomponent')  
    # Relation avec Component (un composant peut avoir plusieurs sous-composants)
    
    class Meta:
        db_table = 'list_subcomponent'
        ordering = ['subcomponent_name']

    def __str__(self):
        return self.subcomponent_name


# ============================
# TABLE DES INDICATEURS
# ============================
class Indicator(models.Model):
    # Choix prédéfinis pour certaines colonnes
    CATOGORY_CHOICES = (
        ('Prevalence', 'Prevalence'),
        ('Incidence', 'Incidence'),
        ('Death', 'Death'),
        ('Strategy', 'Strategy')
    )
    MODELING_CHOICES = (
        ('sir', 'SIR'),
        ('arima', 'ARIMA'),
        ('ACP', 'ACP'),
        ('ACM', 'ACM'),
        ('None', 'None')
    )
    PERFORMANCE_CHOICES = (
        ('Low', 'Low'),
        ('High', 'High'),
        ('Constant', 'Constant')
    )
    TYPE_CHOICES = (
        ('Quanti_ind', 'Quantitatif'),
        ('Quali_ind', 'Qualitatif')
    )

    indicator_code = models.CharField('Indicator code', max_length=30)  # Code de l’indicateur
    indicator_name = models.CharField('Indicator name', max_length=250) # Nom de l’indicateur
    indicator_description = models.CharField('Indicator description', max_length=250, null=True)  # Description
    indicator_target = models.CharField('Target', max_length=250, null=True)                      # Objectif
    indicator_metric = models.CharField('Metric', max_length=20, null=True)                       # Métrique
    indicator_unit = models.CharField('Unit of measurement', max_length=50, null=True)            # Unité
    indicator_frequency = models.CharField('Indicator frequency', max_length=250, null=True)      # Fréquence
    indicator_source = models.URLField('Link', max_length=500)                                    # Source
    subcomponent = models.ForeignKey(Subcomponent, on_delete=models.SET_NULL, null=True, related_name='subcomponent_indicator')  
    # Relation avec Subcomponent
    category_indicator = models.CharField('Category', max_length=10, null=True, choices=CATOGORY_CHOICES)
    forecasting_indicator = models.CharField('Forecasting', max_length=10, null=True, choices=MODELING_CHOICES)
    performance_indicator = models.CharField('Performance', max_length=10, null=True, choices=PERFORMANCE_CHOICES)
    type_indicator = models.CharField('Type of indicator', max_length=15, null=True, choices=TYPE_CHOICES)
    ref_data = models.CharField('Reference of data', max_length=50)  # Référence des données

    class Meta:
        db_table = 'list_indicator'

    def __str__(self):
        return f"{self.subcomponent} {self.indicator_code}"


# ============================
# TABLE DE RÉPOSITORY INDICATEUR (API GHO)
# ============================
class RepositoryIndicator(models.Model):
    indicator = models.ForeignKey(Indicator, on_delete=models.SET_NULL, null=True, related_name='indicator_repository')
    country = models.ForeignKey('Country', on_delete=models.SET_NULL, null=True, related_name='country_repository')
    # Relation avec Indicator et Country
    indicator_code = models.CharField('Indicator code', max_length=250, null=True, blank=True)
    spatial_dim = models.CharField('Spatial Dim', max_length=250, null=True, blank=True)
    dim1_type = models.CharField('Dim1 type', max_length=250, null=True, blank=True)
    dim1 = models.CharField('Dim1', max_length=250, null=True, blank=True)
    dim2_type = models.CharField('Dim2 type', max_length=250, null=True, blank=True)
    dim2 = models.CharField('Dim2', max_length=250, null=True, blank=True)
    dim3_type = models.CharField('Dim3 type', max_length=250, null=True, blank=True)
    dim3 = models.CharField('Dim3', max_length=250, null=True, blank=True)
    time_dim = models.FloatField('Time', null=True, blank=True)  # Dimension temporelle
    alpha_value = models.CharField('Alpha-numeric value', max_length=1000, null=True, blank=True)
    numeric_value = models.DecimalField('Numerical value', max_digits=12, decimal_places=6, null=True, blank=True)
    publish_date = models.DateField('Date of publication', blank=True)

    class Meta:
        db_table = 'data_loader_repository'

    def __str__(self):
        return f"{self.indicator_code} {self.spatial_dim}"
    

    # ============================
# TABLE DES PROJETS D’ENQUÊTE
# ============================
class SurveyProject(models.Model):
    responsible = models.CharField('Organisation responsable', max_length=500, null=True, blank=True)  # Organisation responsable
    title_surv = models.CharField('Sujet de l’enquête', max_length=350, null=True, blank=True)        # Titre / sujet
    target_population = models.CharField('Population cible', max_length=350, null=True, blank=True)   # Population cible
    start_date = models.DateField('Date de début', blank=True)                                        # Date de début
    end_date = models.DateField('Date de fin', blank=True)                                            # Date de fin
    location_survey = models.CharField('Lieu de l’enquête', max_length=500, null=True, blank=True)    # Lieu
    date_creation = models.DateTimeField(auto_now_add=True)                                           # Date de création automatique
    
    class Meta:
        db_table = 'survey_project'

    def __str__(self):
        return self.title_surv


# ============================
# TABLE DES DONNÉES D’ENQUÊTE
# ============================
class SurveyDataset(models.Model):
    surveyProject = models.ForeignKey(SurveyProject, on_delete=models.SET_NULL, null=True, related_name='project_surveyData')  
    # Relation avec SurveyProject
    quest_code = models.CharField('Code question', max_length=250, null=True, blank=True)             # Code de la question
    question = models.CharField('Question', max_length=1000)                                          # Texte de la question
    response_text = models.CharField('Réponse texte', max_length=1000, null=True, blank=True)         # Réponse alphanumérique
    response_num = models.FloatField('Réponse numérique', null=True, blank=True)                      # Réponse numérique
    level_1 = models.CharField('Dimension spatiale 1', max_length=500, null=True, blank=True)         # Niveau spatial 1
    level_2 = models.CharField('Dimension spatiale 2', max_length=500, null=True, blank=True)         # Niveau spatial 2
    
    class Meta:
        db_table = 'data_loader_survey'

    def __str__(self):
        return f"{self.quest_code} {self.question}"


# ============================
# TABLE POUR SAUVEGARDE D’IMAGES
# ============================
class DocSave(models.Model):
    name_doc = models.CharField('Nom du fichier', max_length=200, null=True, blank=True)              # Nom du document
    file_doc = models.ImageField(upload_to='static/images', null=True)                                # Image uploadée
    
    class Meta:
        db_table = 'picture_loader_cover'

    def __str__(self):
        return self.name_doc


# ============================
# TABLE POUR SAUVEGARDE DE RAPPORTS PDF
# ============================
class ReportSave(models.Model):
    title_rep = models.CharField('Titre du rapport', max_length=200, null=True, blank=True)           # Titre
    author_rep = models.CharField('Organisation responsable', max_length=500, null=True, blank=True)  # Auteur / organisation
    date_rep = models.DateField('Date', blank=True)                                                   # Date du rapport
    summary_rep = models.CharField('Résumé', max_length=500)                                          # Résumé
    file_rep = models.FileField(upload_to='static/report/pdf', null=True)                             # Fichier PDF
    img_cp_rep = models.ImageField(upload_to='static/report/picture', null=True)                      # Image associée
    
    class Meta:
        db_table = 'report_loader'

    def __str__(self):
        return self.title_rep


# ============================
# TABLE DES SCRIPTS WAREHOUSE
# ============================
class WarehouseScript(models.Model):
    COLLECT_CHOICES = (
        ('E', 'Empirique'),
        ('S', 'Surveillance')
    )
    ws_title = models.CharField('Nom enquête', max_length=250, null=True)                             # Nom de l’enquête
    ws_type = models.CharField('Type de collecte', max_length=20, choices=COLLECT_CHOICES)            # Type de collecte
    ws_table = models.CharField('Nom table', max_length=250, null=True)                               # Nom de la table
    ws_create = models.CharField('Script CREATE', max_length=500, null=True)                          # Script de création
    ws_insert = models.CharField('Script INSERT', max_length=500, null=True)                          # Script d’insertion
    ws_select = models.CharField('Script SELECT', max_length=500, null=True)                          # Script de sélection
    ws_join = models.CharField('Script JOIN', max_length=500, null=True)                              # Script de jointure
    ws_pivot = models.CharField('Script PIVOT', max_length=500, null=True)                            # Script de pivot
    ws_pathfile = models.CharField('Chemin fichier', max_length=500, null=True)                       # Chemin du fichier
    ws_focalpoint = models.CharField('Point focal', max_length=200, null=True)                        # Responsable
    ws_date = models.DateTimeField(auto_now_add=True)                                                 # Date de création
    
    class Meta:
        db_table = 'scripts_warehouse'

    def __str__(self):
        return self.ws_table



# ============================
# EXTRACTION DE DONNÉES API (ApiExtract)
# ============================



class ApiExtract(models.Model):
    indicator_code = models.CharField(max_length=100, blank=True, null=True)
    indicator_name = models.CharField(max_length=250, blank=True, null=True) # Nom de l’indicateur
    disease_name = models.CharField(max_length=250, blank=True, null=True) # Nom de la maladie
    country_code = models.CharField(max_length=50, blank=True, null=True)
    country_name = models.CharField(max_length=250, blank=True, null=True) # Nom de l’indicateur
    parent_location = models.CharField(max_length=100, blank=True, null=True)
    time_dim = models.IntegerField(blank=True, null=True)
    Dim1 = models.CharField(max_length=250, blank=True, null=True)
    Dim2 = models.CharField(max_length=250, blank=True, null=True)
    Dim3 = models.CharField(max_length=250, blank=True, null=True)     
    alpha_value = models.CharField(max_length=255, blank=True, null=True)
    numeric_value = models.FloatField(blank=True, null=True)
    date = models.DateField(blank=True, null=True)
    api_id = models.CharField(max_length=255, unique=True)

    class Meta:
        db_table = "data_loader_api"

    def __str__(self):
        return f"{self.indicator_code} - {self.country_code} ({self.date})"


class ApiFieldConfig(models.Model):
    field_name = models.CharField(max_length=200, unique=True)
    enabled = models.BooleanField(default=True)

    class Meta:
        db_table = "api_field_config"
        verbose_name = "API Field Configuration"
        verbose_name_plural = "API Field Configurations"

    def __str__(self):
        return f"{self.field_name} ({'enabled' if self.enabled else 'disabled'})"


# ============================
# DONNÉES API (ApiData)
# ============================
class ApiData(models.Model):
    api_id = models.IntegerField(unique=True)                                  # Identifiant API
    IndicatorCode = models.CharField(max_length=255, null=True, blank=True)    # Code indicateur
    SpatialDimType = models.CharField(max_length=255, null=True, blank=True)   # Type dimension spatiale
    SpatialDim = models.CharField(max_length=255, null=True, blank=True)       # Dimension spatiale
    TimeDimType = models.CharField(max_length=255, null=True, blank=True)      # Type dimension temporelle
    ParentLocationCode = models.CharField(max_length=255, null=True, blank=True) # Code lieu parent
    ParentLocation = models.CharField(max_length=255, null=True, blank=True)   # Nom lieu parent
    Dim1Type = models.CharField(max_length=255, null=True, blank=True)         # Type dimension 1
    Dim1 = models.CharField(max_length=255, null=True, blank=True)             # Valeur dimension 1
    TimeDim = models.FloatField(null=True, blank=True)                         # Dimension temporelle
    Dim2Type = models.CharField(max_length=255, blank=True, null=True)         # Type dimension 2
    Dim2 = models.CharField(max_length=255, blank=True, null=True)             # Valeur dimension 2
    Dim3Type = models.CharField(max_length=255, blank=True, null=True)         # Type dimension 3
    Dim3 = models.CharField(max_length=255, blank=True, null=True)             # Valeur dimension 3
    DataSourceDimType = models.CharField(max_length=255, blank=True, null=True)# Type source
    DataSourceDim = models.CharField(max_length=255, blank=True, null=True)    # Source
    Value = models.CharField(max_length=255, blank=True, null=True)            # Valeur brute
    NumericValue = models.DecimalField(max_digits=12, decimal_places=12, null=True, blank=True) # Valeur numérique
    Low = models.CharField(max_length=255, blank=True, null=True)              # Valeur basse
    High = models.CharField(max_length=255, blank=True, null=True)             # Valeur haute
    Comments = models.CharField(max_length=255, blank=True, null=True)         # Commentaires
    Date = models.CharField(max_length=255)                                    # Date
    TimeDimensionValue = models.CharField(max_length=255, blank=True, null=True) # Valeur temporelle
    TimeDimensionBegin = models.CharField(max_length=255, blank=True, null=True) # Début période
    TimeDimensionEnd = models.CharField(max_length=255, blank=True, null=True)   # Fin période
    created_at = models.DateTimeField(auto_now_add=True)                       # Date création
    updated_at = models.DateTimeField(auto_now=True)                           # Date mise à jour
    
    class Meta:
        db_table = 'dataset_loader_api_new'

    def __str__(self):
        return f"{self.IndicatorCode} - {self.SpatialDim} ({self.TimeDimensionValue})"


# ============================
# STOCKAGE DES DONNÉES API (StoreAPI)
# ============================
class StoreAPI(models.Model):
    api_id = models.IntegerField(unique=True)                                  # Identifiant API
    indicator_code = models.CharField('Indicator code', max_length=255, null=True, blank=True) # Code indicateur
    country_code = models.CharField('Country code', max_length=255, null=True, blank=True)     # Code pays
    dim1_type = models.CharField(max_length=255, null=True, blank=True)        # Type dimension 1
    dim1 = models.CharField(max_length=255, null=True, blank=True)             # Valeur dimension 1
    dim2_type = models.CharField(max_length=255, blank=True, null=True)        # Type dimension 2
    dim2 = models.CharField(max_length=255, blank=True, null=True)             # Valeur dimension 2
    dim3_type = models.CharField(max_length=255, blank=True, null=True)        # Type dimension 3
    dim3 = models.CharField(max_length=255, blank=True, null=True)             # Valeur dimension 3
    time_dim = models.FloatField(null=True, blank=True)                        # Dimension temporelle
    alpha_value = models.CharField(max_length=255, blank=True, null=True)      # Valeur alphanumérique
    numeric_value = models.FloatField(null=True, blank=True)                   # Valeur numérique
    publish_date = models.CharField(max_length=255)                            # Date publication
    created_at = models.DateTimeField(auto_now_add=True)                       # Date
 
    class Meta:
        db_table = 'store_api'
       
    def __str__(self):
        return f"{self.indicator_code} - {self.country_code} ({self.time_dim})"
    



class ApiFieldMapping(models.Model):
    api_field = models.CharField(max_length=100)
    db_field = models.CharField(max_length=100)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        db_table = 'mapping_api_db'
       

    def __str__(self):
        return f"{self.api_field} → {self.db_field}"

 


class StaffMember(models.Model):
   
    country = models.ManyToManyField("Country", blank=True)
    grade = models.CharField("Grade", max_length=50, blank=True, null=True)
    name = models.CharField("Full Name", max_length=150)
    email = models.EmailField("Email", unique=True)
    telephone = models.CharField("Telephone", max_length=30, blank=True, null=True)
    office_affiliation = models.CharField("Office Affiliation", max_length=150, blank=True, null=True)
    position = models.CharField("Position", max_length=100, blank=True, null=True)
    responsibility = models.TextField("Responsibility", blank=True, null=True)
    language = models.JSONField("Languages", blank=True, null=True)  # stockage liste
    diseases = models.ManyToManyField("Subcomponent", related_name="staff_members", blank=True)
    level_geo = models.CharField("Geographical Level", max_length=50, blank=True, null=True)

    class Meta:
        db_table = "staff_member"
        ordering = ["name"]

    def __str__(self):
        return f"{self.name}"


