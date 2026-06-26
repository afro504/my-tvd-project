# =========================
# DJANGO CORE IMPORTS
# =========================
import io
from django.shortcuts import render, get_object_or_404, redirect
from django.urls import reverse
from django.http import HttpResponse, JsonResponse, FileResponse, HttpResponseRedirect
from django.core.paginator import Paginator, PageNotAnInteger, EmptyPage
from django.db import transaction, IntegrityError
from django.db.models import (
    Count, Avg, Min, Max, Sum,
    CharField, Value, Q, Case, When, F, OuterRef, BooleanField
)
from django.db.models import Prefetch
from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.models import User
from django.contrib.auth.decorators import login_required, permission_required, user_passes_test
from django.views.decorators.csrf import csrf_exempt
from docx import Document

# =========================
# DJANGO GENERIC VIEWS
# =========================
from django.views.generic import TemplateView, ListView, CreateView, DetailView, FormView
from django.views.generic.detail import SingleObjectMixin
 
# =========================
# DJANGO FORMS
# =========================
from django.forms.models import (
    inlineformset_factory,
    formset_factory,
    modelform_factory,
    modelformset_factory
)
 
# =========================
# PROJECT MODELS
# =========================
from .models import (
    Country, Component, Subcomponent, Indicator,
    RepositoryIndicator, SurveyProject, SurveyDataset,
    DocSave, ReportSave, LocationCountry, City,
    WarehouseScript, ApiData, StoreAPI,ApiExtract, 
    ApiFieldConfig,ApiFieldMapping,StaffMember
)



# =========================
# PROJECT FORMS
# =========================
from .forms import (
    CountryForm, IndicatorForm, IndicatorSearchForm,
    ComponentForm, Loadcomponent, SelectcomponentForm,
    ComponentsForm, SubcomponentsForm, SelectSubcomponentForm,
    DocSaveForm, ReportSaveForm, SurveyProjectForm,
    SelectUrlForm, SurveyDatasetForm, SelectSurveyForm,
    RepositoryIndicatorForm, WarehouseScriptForm,SelectUrlForms,
    S_table_nameForm, RepositoryUploadForm,
    SelectRepositoryForm, 
    SelectAPIForm, StoreAPIForm,  ApiFieldSelectionForm,StaffMemberForm
)
 
# =========================
# EXTERNAL LIBRARIES
# =========================
import pandas as pd
from datetime import datetime
import json
import requests
import folium
import streamlit as st
from streamlit_folium import st_folium
import csv

 
# =========================
# EXCEL / DATA EXPORT
# =========================
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Color, PatternFill, Font, Border, Alignment, Side
from tablib import Dataset
from django.http import HttpResponse
from openpyxl.utils import get_column_letter
 
# =========================
# EMAIL / AUTH UTILITIES
# =========================
from django.core.mail import send_mail
from django.template.loader import render_to_string, get_template
from django.contrib.sites.shortcuts import get_current_site
from django.utils.http import urlsafe_base64_decode, urlsafe_base64_encode
from django.utils.encoding import force_bytes, force_str
from django.core.mail import send_mail
from django.contrib import messages
 
# Compatibilité ancienne version Django
try:
    from django.utils.encoding import force_text
except ImportError:
    from django.utils.encoding import force_str as force_text
 
# =========================
# PROJECT SETTINGS
# =========================
from TVDdata import settings
 
# =========================
# REST FRAMEWORK
# =========================
from rest_framework.response import Response
 
# =========================
# PROJECT RESOURCES
# =========================
from .resources import SurveyResource
import re




# =========================
# COVER PAGE WITH MAP
# =========================
def coverPage(request):
    """
    Affiche une carte Folium avec les localisations depuis la base de données
    """
 
    # ✅ Création de la carte
    m = folium.Map(
        location=(-8.783195, 34.508523),
        zoom_start=3,
        tiles="cartodb positron",
        scrollWheelZoom=True
    )
 
    folium.LayerControl().add_to(m)
 
    # ✅ Récupération des données (optimisée)
    locations = LocationCountry.objects.only('name', 'latitude', 'longitude')
 
    # ✅ Ajout des markers
    for location in locations:
        if location.latitude and location.longitude:
            folium.Marker(
                location=(location.latitude, location.longitude),
                popup=f"<b>{location.name}</b><br>Lat: {location.latitude}, Lon: {location.longitude}",
                tooltip=location.name
            ).add_to(m)
 
    # ✅ Conversion en HTML (IMPORTANT : en dehors de la boucle)
    map_html = m._repr_html_()
 
    context = {
        'm': map_html
    }
 
    return render(request, 'mytvddata/cover/index.html', context)

 
# =========================
# SIGNUP VIEW
# =========================
def signup(request):
    if request.method == "POST":
        username = request.POST.get('username')
        fname = request.POST.get('fname')
        lname = request.POST.get('lname')
        email = request.POST.get('email')
        pass1 = request.POST.get('pass1')
        pass2 = request.POST.get('pass2')
 
        # ✅ Vérifier si username existe
        if User.objects.filter(username=username).exists():
            messages.error(request, "Username already exists")
            return render(request, 'mytvddata/cover/signup.html')
 
        # ✅ Vérifier email existe déjà
        if User.objects.filter(email=email).exists():
            messages.error(request, "Email already registered")
            return render(request, 'mytvddata/cover/signup.html')
 
        # ✅ ✅ VALIDATION EMAIL (format)
        email_regex = r'^[\w\.-]+@[\w\.-]+\.\w+$'
        if not re.match(email_regex, email):
            messages.error(request, "Invalid email format")
            return render(request, 'mytvddata/cover/signup.html')
 
        # ✅ Longueur username
        if len(username) > 15:
            messages.error(request, "Username must be under 15 characters")
            return render(request, 'mytvddata/cover/signup.html')
 
        # ✅ Vérifier caractères username
        if not username.isalnum():
            messages.error(request, "Username must be alphanumeric")
            return render(request, 'mytvddata/cover/signup.html')
 
        # ✅ Vérifier mots de passe identiques
        if pass1 != pass2:
            messages.error(request, "Passwords do not match")
            return render(request, 'mytvddata/cover/signup.html')
 
        # ✅ ✅ SÉCURITÉ MOT DE PASSE
        if len(pass1) < 8:
            messages.error(request, "Password must be at least 8 characters long")
            return render(request, 'mytvddata/cover/signup.html')
 
        if not re.search(r'[A-Z]', pass1):
            messages.error(request, "Password must contain at least one uppercase letter")
            return render(request, 'mytvddata/cover/signup.html')
 
        if not re.search(r'[a-z]', pass1):
            messages.error(request, "Password must contain at least one lowercase letter")
            return render(request, 'mytvddata/cover/signup.html')
 
        if not re.search(r'[0-9]', pass1):
            messages.error(request, "Password must contain at least one number")
            return render(request, 'mytvddata/cover/signup.html')
 
        # ✅ Création utilisateur
        user = User.objects.create_user(
            username=username,
            email=email,
            password=pass1
        )
 
        user.first_name = fname
        user.last_name = lname
        user.is_active = True
 
        user.save()
 
        messages.success(request, "Account created successfully")
        return redirect('signin')
 
    return render(request, 'mytvddata/cover/signup.html')
 
 
# =========================
# SIGNIN VIEW
# =========================
def signin(request):
    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('pass1')
 
        user = authenticate(request, username=username, password=password)
 
        if user is not None:
            login(request, user)
            messages.success(request, "Logged in successfully")
            return redirect('coverPage')
        else:
            messages.error(request, "Invalid username or password")
            return render(request, 'mytvddata/cover/signin.html')
 
    return render(request, 'mytvddata/cover/signin.html')
 
# =========================
# SIGNOUT VIEW
# =========================
@login_required
def signout(request):
    logout(request)
    messages.success(request, "Logged out successfully")
    return redirect('coverPage')
 
 
def countries(request):
    country_list = Country.objects.order_by('name')
 
    query = request.GET.get('q')
    if query:
        country_list = country_list.filter(name__icontains=query)
 
    paginator = Paginator(country_list, 10)
    countries_page = paginator.get_page(request.GET.get('page'))
 
    return render(request, 'mytvddata/pages/country/countries.html', {
        'country_list': countries_page,
        'total_country': paginator.count
    })


 
# =========================
# CRUD COUNTRY (PRO VERSION)
# =========================
@login_required
def add_country(request):
 
    # ✅ Base queryset optimisé
    country_list = Country.objects.all().order_by('name')
 
    # =========================
    # SEARCH
    # =========================
    search_query = request.GET.get('name')
    if search_query:
        country_list = country_list.filter(name__icontains=search_query)
 
    # =========================
    # PAGINATION (VERSION PRO)
    # =========================
    paginator = Paginator(country_list, 50)
    countries_page = paginator.get_page(request.GET.get('page'))
 
    # =========================
    # FORM INIT
    # =========================
    form = CountryForm()
 
    # =========================
    # POST ACTIONS (CRUD)
    # =========================
    if request.method == 'POST':
 
        # ✅ SAVE / UPDATE
        if 'save' in request.POST:
            pk = request.POST.get('save')
 
            if pk:  # update
                country = get_object_or_404(Country, id=pk)
                form = CountryForm(request.POST, instance=country)
            else:  # create
                form = CountryForm(request.POST)
 
            if form.is_valid():
                form.save()
                messages.success(request, "Country saved successfully")
                return redirect('add_country')  # ✅ éviter resubmit
 
        # ✅ DELETE
        elif 'delete' in request.POST:
            if request.user.groups.filter(name='Manager and delete').exists():
 
                pk = request.POST.get('delete')
                country = get_object_or_404(Country, id=pk)
                country.delete()
 
                messages.success(request, "Country deleted successfully")
                return redirect('add_country')
 
            else:
                return render(request, 'pages/examples/403.html')
 
        # ✅ EDIT (pré-remplir form)
        elif 'edit' in request.POST:
            pk = request.POST.get('edit')
            country = get_object_or_404(Country, id=pk)
            form = CountryForm(instance=country)
 
    # =========================
    # CONTEXT
    # =========================
    context = {
        'countries': countries_page,
        'total_country': paginator.count,
        'form': form,
        'search_query': search_query
    }
 
    return render(request, 'mytvddata/pages/country/add_country.html', context)
 

# =========================
# EXPORT COUNTRY PRO MAX
# =========================
@login_required
def export_country(request):
 
    # ✅ Réponse HTTP
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=Country_Report.xlsx'
 
    wb = Workbook()
 
    # =========================
    # SHEET 1: ALL COUNTRIES
    # =========================
    ws = wb.active
    ws.title = "Countries"
 
    headers = [
        "ID", "Name", "CCA3", "Capital",
        "Subregion", "Population"
    ]
 
    ws.append(headers)
 
    # ✅ STYLE HEADER
    header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    header_font = Font(bold=True)
 
    for col_num, col_title in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
 
    # =========================
    # FILTRE DYNAMIQUE
    # =========================
    query = request.GET.get('name')
    countries = Country.objects.all()
 
    if query:
        countries = countries.filter(name__icontains=query)
 
    # =========================
    # DATA
    # =========================
    for c in countries:
        ws.append([
            c.id,
            c.name,
            c.cca3,
            c.capital,
            c.subregion,
            c.population
        ])
 
    # ✅ FILTER EXCEL
    ws.auto_filter.ref = ws.dimensions
 
    # ✅ AUTO WIDTH
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
 
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
 
        ws.column_dimensions[col_letter].width = max_length + 2
 
    # =========================
    # SHEET 2: SUMMARY
    # =========================
    ws2 = wb.create_sheet(title="Summary")
 
    ws2.append(["Metric", "Value"])
    ws2.append(["Total Countries", countries.count()])
 
    # Exemple métrique supplémentaire
    total_pop = sum([c.population or 0 for c in countries])
    ws2.append(["Total Population", total_pop])
 
    # =========================
    # SAVE
    # =========================
    wb.save(response)
 
    return response
 

 
@login_required
def list_indicator(request):
 
    # ✅ base queryset optimisé
    indicators_qs = Indicator.objects.select_related('subcomponent').order_by('indicator_name')
 
    form_subcomponent = SelectSubcomponentForm(request.GET or None)
 
    # ✅ filtres
    search_name = request.GET.get('indicator_name')
    by_subcomponent = request.GET.get('by_subcomponent')
 
    if search_name:
        indicators_qs = indicators_qs.filter(indicator_name__icontains=search_name)
 
    if by_subcomponent:
        indicators_qs = indicators_qs.filter(subcomponent__id=by_subcomponent)
 
    # ✅ pagination simplifiée
    paginator = Paginator(indicators_qs, 10)
    indicators_page = paginator.get_page(request.GET.get('page'))
 
    context = {
        'indicators': indicators_page,
        'total_indicator': paginator.count,
        'form_subcomponent': form_subcomponent,
        'by_subcomponent': by_subcomponent,
    }
 
    return render(request, 'mytvddata/pages/indicator/indicator_list.html', context)

 


 
def save_indicator_form(request, form, template_name):
    data = {}
 
    if request.method == 'POST':
        if form.is_valid():
            form.save()
            data['form_is_valid'] = True
 
            indicators = Indicator.objects.select_related('subcomponent')
 
            data['html_indicator_list'] = render_to_string(
                'mytvddata/pages/indicator/partial_indicator_list.html',
                {'indicators': indicators},
                request=request
            )
        else:
            data['form_is_valid'] = False
 
    # ✅ toujours renvoyer le formulaire
    context = {'form': form}
    data['html_form'] = render_to_string(template_name, context, request=request)
 
    return JsonResponse(data)
 

@login_required
def create_indicator(request):
    if request.method == 'POST':
        form = IndicatorForm(request.POST)
 
        if form.is_valid():
            messages.success(request, 'Indicator saved successfully')
 
    else:
        form = IndicatorForm()
 
    return save_indicator_form(
        request,
        form,
        'mytvddata/pages/indicator/partial_indicator_create.html'
    )
 

@login_required
def update_indicator(request, pk):
 
    indicator = get_object_or_404(Indicator, pk=pk)
 
    if request.method == 'POST':
        form = IndicatorForm(request.POST, instance=indicator)
 
        if form.is_valid():
            messages.success(request, 'Indicator updated successfully')
    else:
        form = IndicatorForm(instance=indicator)
 
    return save_indicator_form(
        request,
        form,
        'mytvddata/pages/indicator/partial_indicator_update.html'
    )
 

@login_required
def delete_indicator(request, pk):
 
    indicator = get_object_or_404(Indicator, pk=pk)
    data = {}
 
    if request.method == 'POST':
        indicator.delete()
        data['form_is_valid'] = True
 
        indicators = Indicator.objects.select_related('subcomponent')
 
        data['html_indicator_list'] = render_to_string(
            'mytvddata/pages/indicator/partial_indicator_list.html',
            {'indicators': indicators},
            request=request
        )
 
        messages.success(request, 'Indicator deleted successfully')
 
    else:
        context = {'indicator': indicator}
        data['html_form'] = render_to_string(
            'mytvddata/pages/indicator/partial_indicator_delete.html',
            context,
            request=request
        )
 
    return JsonResponse(data)
 


 
@login_required
def export_indicator(request):
 
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=Indicator.xlsx'
 
    wb = Workbook()
    ws = wb.active
    ws.title = "Indicators"
 
    # ✅ headers propres
    headers = [
        "ID", "Code", "Name", "Description",
        "Target", "Metric", "Unit",
        "Frequency", "Type", "Source",
        "Subcomponent", "Category",
        "Forecast", "Performance", "Ref Data"
    ]
    ws.append(headers)
 
    # ✅ queryset optimisé
    indicators = Indicator.objects.select_related('subcomponent')
 
    for i in indicators:
        ws.append([
            i.id,
            i.indicator_code,
            i.indicator_name,
            i.indicator_description,
            i.indicator_target,
            i.indicator_metric,
            i.indicator_unit,
            i.indicator_frequency,
            i.type_indicator,
            i.indicator_source,
            i.subcomponent.subcomponent_name if i.subcomponent else '',
            i.category_indicator,
            i.forecasting_indicator,
            i.performance_indicator,
            i.ref_data
        ])
 
    wb.save(response)
    return response
 



# CREATE VIEWS FOR SCRIPTS TABLE------------------------------------

def t_scripts_view(request):
  
    form_survey_title = S_table_nameForm(request.GET or None)
    by_title =  request.GET.get('by_title')
    end_date =  request.GET.get('end')
  
    t_scripts_views = WarehouseScript.objects.all().order_by("ws_title")
    form = WarehouseScriptForm()
    if request.method=='GET':
        reseash_s_focal_point=request.GET.get('ws_focalpoint')
        if reseash_s_focal_point!=None:
            t_scripts_views= WarehouseScript.objects.filter(ws_focalpoint__icontains=reseash_s_focal_point)
        if by_title!=None:
            t_scripts_views = WarehouseScript.objects.all().filter(Q(id__icontains=by_title),Q(ws_date__lte=end_date)).order_by("ws_title")# id use to filter the title because the paramters is id but not label
               
    context={'form_survey_title': form_survey_title,
             'form':form,
             't_scripts_views': t_scripts_views,
             'by_title':by_title,
             'end_date':end_date
             }
        
    return render(request, 'mytvddata/pages/scripts_view/t_scripts_view.html', context)

# Method for delete SCRIPTS

def delete_t_scripts(request,pk):
    data=WarehouseScript.objects.filter(id=pk)
    data.delete()
    messages.warning(request, 'Scripts deleted successfully')
   
    return redirect('t_scripts_view')


# Method for add new SCRIPTS

def add_new_t_scripts(request):
    if request.method == "POST":
        form = WarehouseScriptForm(request.POST)
        if form.is_valid():
            form.save(commit=True)
          #  form.save_m2m
         #   return JsonResponse({'msg': 'Data saved'})
            messages.success(request, "SCRIPTS Added Successfully")
            return redirect('t_scripts_view')
        else:
            print("ERROR FORM INVALID")
            return JsonResponse({'msg': 'ERROR FORM INVALID'})
    else:
   #     form = WorkplanForm()
  #  return JsonResponse({'form': form})
        context = {
            'form': form
        }
    return render(request, 'mytvddata/pages/scripts_view/add_t_scripts_modal.html', context)

         
# Method for edit SCRIPTS
def edit_t_scripts(request, pk): 
    t_script = WarehouseScript.objects.get(id=pk)
    form = WarehouseScriptForm(instance=t_script)
    
    if request.method == "POST":
        form = WarehouseScriptForm(request.POST, instance=t_script)
        if form.is_valid():
            form.save()
            messages.success(request, "SCRIPTS updated Successfully")
            return redirect('t_scripts_view')
    messages.warning(request, "ERROR FORM INVALID")
  
    return render(request, 'mytvddata/pages/scripts_view/t_scripts_view.html', {'form': form})

# Method for import and export Table SurveyDataset
def survey_add_project(request):
    context={}
    form = SurveyProjectForm()
    surveyProjects = SurveyProject.objects.all()
    
    if request.method=='GET':
        st=request.GET.get('title_surv')
        if st!=None:
            surveyProjects= SurveyProject.objects.filter(title_surv__icontains=st)
    
    page = request.GET.get('page')
    num_of_items = 10
    # paginator of country 
    paginator = Paginator(surveyProjects, num_of_items)
       
    try:
        surveyProjects = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        surveyProjects = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        surveyProjects = paginator.page(page)
            
    context['surveyProjects'] = surveyProjects
    context['paginator']=paginator
    context['title'] ='home'
    if request.method == 'POST':
        if 'save' in request.POST:
            pk = request.POST.get('save')
            if not pk:
                form=SurveyProjectForm(request.POST)
            else:
                surveyProject = SurveyProject.objects.get(id = pk)
                form = SurveyProjectForm(request.POST ,instance = surveyProject) 
            form.save()
            form = SurveyProjectForm()
            messages.success(request, 'Survey project saved or updated successfully')
        elif 'delete' in request.POST:
            if request.user.groups.filter(name='Manager and delete').exists(): ## Condition security
                pk = request.POST.get('delete')
                surveyProject = SurveyProject.objects.get(id = pk)
                surveyProject.delete()
                messages.success(request, 'Survey project deleted successfully')
            else:
                return render(request, 'pages/examples/403.html')
    
        elif 'edit' in request.POST:
            pk = request.POST.get('edit')
            surveyProject = SurveyProject.objects.get(id = pk)
            form = SurveyProjectForm(instance = surveyProject)
          
    context['form'] = form
   
    return render(request, 'mytvddata/pages/survey/survey_project.html', context)


# METHOD FOR IMPORT AND EXPORT TABLE SURVEYDATASET
def simple_upload(request):
    if request.method =='POST':
        survey_resource = SurveyResource()
        dataset = Dataset()
        new_survey = request.FILES['myfile']
        
        if not new_survey.name.endswith('xlsx'):
            messages.info(request, 'wrong format')
            return render(request, 'mytvddata/pages/survey/upload.html')
        
        import_data = dataset.load(new_survey.read(),format='xlsx')
        for data in import_data:
            value = SurveyDataset(
                data[0],
                data[1],
                data[2],
                data[3],
                data[4],
                data[5],
                data[6],
                data[7]
            )
            value.save()
    return render(request, 'mytvddata/pages/survey/upload.html')


# Method to add new data into the SuverDataset
#@login_required
#@group_required('Technical officer')
def survey_add_data(request):
  #  ct=ContentType.objects.get_for_model(SurveyDataset)
  #  if request.user.permissions.filter(codename="viw_surveydataset", contentype=ct).exists():
        
    context={}
    form_survey_title = SelectSurveyForm(request.GET or None)
    by_survey =  request.GET.get('by_survey')
    end_day =  request.GET.get('end')
 
    form = SurveyDatasetForm()
    surveyDatasets = SurveyDataset.objects.all().order_by("-surveyProject__end_date")
    if request.method=='GET':
        st=request.GET.get('quest_code')
        if st!=None:
            surveyDatasets= SurveyDataset.objects.filter(quest_code__icontains=st).order_by("-surveyProject__end_date")
            
        if by_survey!=None:
            surveyDatasets = SurveyDataset.objects.filter(Q(surveyProject__id__icontains=by_survey)& Q(surveyProject__end_date__lte=end_day)).order_by("quest_code")
            
    #page = request.GET.get('page')
    #num_of_items = 10
  
  #  paginator = Paginator(surveyDatasets, num_of_items)
       
  #  try:
  #      surveyDatasets = paginator.page(page)
  #  except PageNotAnInteger:
  #      page = 1
  #      surveyDatasets = paginator.page(page)
  #  except EmptyPage:
  #      page = paginator.num_pages
  #      surveyDatasets = paginator.page(page)
    context['surveyDatasets'] = surveyDatasets
   # context['paginator']=paginator
    context['title'] ='home'
  #  if request.user.groups.filter(name='Technical officer').exists():
    
    if request.method == 'POST':
            if 'save' in request.POST:
                pk = request.POST.get('save')  
                if not pk:
                    form=SurveyDatasetForm(request.POST)
                else:
                    surveyDataset = SurveyDataset.objects.get(id = pk)
                    form = SurveyDatasetForm(request.POST ,instance = surveyDataset) 
                form.save()
                form = SurveyDatasetForm()
                messages.success(request, 'Survey dataset saved or updated successfully')
      
            elif 'delete' in request.POST:
                if request.user.groups.filter(name='Manager and delete').exists():
                    pk = request.POST.get('delete')
                    surveyDataset = SurveyDataset.objects.get(id = pk)
                    surveyDataset.delete()
                    messages.success(request, 'Survey dataset deleted successfully')
                else:
                    return render(request, 'pages/examples/403.html')
            
            elif 'edit' in request.POST:
                pk = request.POST.get('edit')
                surveyDataset = SurveyDataset.objects.get(id = pk)
                form = SurveyDatasetForm(instance = surveyDataset)
            
            context['form'] = form
            context['form_survey_title'] = form_survey_title 
            context['by_survey'] = by_survey
            context['end_day'] = end_day
  #  else:
  #      context['form'] = form
   #     context['form_survey_title'] = form_survey_title 
   #     context['by_survey'] = by_survey
   #     context['end_day'] = end_day
   #     return render(request, 'mytvddata/pages/survey/survey_dataset.html', context)
   
 #   if request.user.has_perm('myuhp.view_surveydataset'):
    return render(request, 'mytvddata/pages/survey/survey_dataset.html', context)


# Method for list all Survey dataset submitted
#@permission_required('myuhp.view_surveydataset',raise_exception=True)

def survey_report(request):
    surveyProjects = SurveyProject.objects.all().order_by("title_surv") 
    if request.method=='GET':
        st=request.GET.get('responsible')
        if st!=None:
            surveyProjects= SurveyProject.objects.filter(responsible__icontains=st)
    
    page = request.GET.get('page')
    num_of_items = 10
    # paginator of country 
    paginator = Paginator(surveyProjects, num_of_items)
       
    try:
        surveyProjects = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        surveyProjects = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        surveyProjects = paginator.page(page)
            
    data={'surveyProjects': surveyProjects, 'paginator':paginator}
    
    return render(request, 'mytvddata/pages/survey/survey_report.html', data)

# Method for view single Survey DataSet page

def single_survey_page(request, pk):
    try:
        surveyProject=SurveyProject.objects.get(id=pk) 
        surveyDatasets =SurveyProject.objects.filter(Q(id__icontains=pk) & Q(project_surveyData__quest_code__isnull=False)).values(
        'responsible',
        'title_surv',
        'start_date',
        'end_date',
        'location_survey',
        'project_surveyData__quest_code',
        'project_surveyData__question',
        'project_surveyData__response_text',
        'project_surveyData__response_num',
        'project_surveyData__level_1',
        'project_surveyData__level_2'
    ).order_by('-start_date')
        
        
        page = request.GET.get('page')
        num_of_items = 10
        # paginator of country 
        paginator = Paginator(surveyDatasets, num_of_items)
        
        try:
            surveyDatasets = paginator.page(page)
        except PageNotAnInteger:
            page = 1
            surveyDatasets = paginator.page(page)
        except EmptyPage:
            page = paginator.num_pages
            surveyDatasets = paginator.page(page)
        
        data={'surveyProject':surveyProject, 'surveyDatasets':surveyDatasets, 'paginator':paginator}   
    except:
        data={'message':'The Survey DataSet you requested doesn t\' exist or isn t in the submitted results'}
    
    return render(request,'mytvddata/pages/survey/single_survey_page.html',data)


#Method to export ALL SURVEY PROJECT on Excel file
def export_project_survey(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Project_survey.xlsx"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Project Survey"

    # Add headers
    headers = ["id","responsible", "title_surv", "start_date", "end_date", "location_survey"]
    ws.append(headers)

    # Add data from the model
    surveyProjects = SurveyProject.objects.all()
    for surveyProject in surveyProjects:
        ws.append([surveyProject.pk, surveyProject.responsible, surveyProject.title_surv, surveyProject.start_date, surveyProject.end_date, surveyProject.location_survey])

    # Save the workbook to the HttpResponse
    wb.save(response)
    return response

#Method to export SURVEY DATASET ALL DATA on Excel file
def export_to_excel(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="survey.xlsx"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Survey"

    # Add headers
    headers = ["responsible", "title_surv", "start_date", "end_date", "location_survey", "quest_code", "question", "response_text", "response_num","level_1","level_2"]
    ws.append(headers)

    # Add data from the model
    surveyDatasets = SurveyDataset.objects.all()
    for s in surveyDatasets:
        ws.append([s.surveyProject.responsible, s.surveyProject.title_surv, s.surveyProject.start_date, s.surveyProject.end_date, s.surveyProject.location_survey, 
                   s.quest_code,  s.question,  s.response_text, s.response_num, s.level_1, s.level_2])

    # Save the workbook to the HttpResponse
    wb.save(response)
    return response

#Method to export DATASET on Excel file ONE BY ONE PROJECT




 
 
# =========================
# EXPORT SURVEY PAGE (PRO)
# =========================
@login_required
def export_survey_page(request, pk):
 
    # ✅ Réponse HTTP moderne (.xlsx)
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=survey_dataset.xlsx'
 
    wb = Workbook()
    ws = wb.active
    ws.title = "Survey"
 
    # =========================
    # HEADERS
    # =========================
    headers = [
        'Responsible', 'Title', 'Start Date', 'End Date',
        'Location', 'Question Code', 'Question',
        'Response Text', 'Response Num', 'Level 1', 'Level 2'
    ]
 
    ws.append(headers)
 
    # =========================
    # QUERY OPTIMISÉE
    # =========================
    projects = SurveyProject.objects.filter(pk=pk).prefetch_related(
        Prefetch(
            'project_surveyData',
            queryset=SurveyDataset.objects.filter(quest_code__isnull=False)
        )
    )
 
    # =========================
    # DATA
    # =========================
    for project in projects:
        for data in project.project_surveyData.all():
            ws.append([
                project.responsible,
                project.title_surv,
                project.start_date,
                project.end_date,
                project.location_survey,
                data.quest_code,
                data.question,
                data.response_text,
                data.response_num,
                data.level_1,
                data.level_2
            ])
 
    # =========================
    # AUTO WIDTH (UX)
    # =========================
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
 
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
 
        ws.column_dimensions[col_letter].width = max_length + 2
 
    # =========================
    # SAVE
    # =========================
    wb.save(response)
 
    return response
 

              
              
## UPLOAD DATASET OF REPOSITORY FORM EXCEL FILE
def upload_repository(request):
    
    if request.method =='POST':
        survey_resource = SurveyResource()
        dataset = Dataset()
        new_repository = request.FILES['myfile']
        
        if not new_repository.name.endswith('xlsx'):
            messages.info(request, 'wrong format')
            return render(request, 'mytvddata/pages/repository/upload_repository.html')
        
        import_data = dataset.load(new_repository.read(),format='xlsx')
        for data in import_data:
            value = RepositoryIndicator(
                data[0],
                data[1],
                data[2],
                data[3],
                data[4],
                data[5],
                data[6],
                data[7],
                data[8],
                data[9],
                data[10],
                data[11],
                data[12],
                data[13],
                data[14]
            )
            value.save()
                  
    return render(request, 'mytvddata/pages/repository/upload_repository.html')



def upload_repository(request):
    preview_data = None
    if request.method == 'POST':
        new_repository = request.FILES['myfile']
        
        if not new_repository.name.endswith('xlsx'):
            messages.info(request, 'Wrong format, please upload .xlsx')
            return render(request, 'mytvddata/pages/repository/upload_repository.html')

        # ✅ Lire avec pandas pour prévisualiser
        df = pd.read_excel(new_repository)
        preview_data = df.head(3).to_html(classes="table table-striped")

        # Stocker le fichier temporairement en session pour mapping
        request.session['uploaded_file'] = new_repository.read()

    return render(request, 'mytvddata/pages/repository/upload_repository.html', {
        'preview_data': preview_data
    })


def repository_add_data(request):
  #  ct=ContentType.objects.get_for_model(SurveyDataset)
  #  if request.user.permissions.filter(codename="viw_repositoryindicator", contentype=ct).exists():
        
    context={}
    form_select_repository= SelectRepositoryForm(request.GET or None)
    by_subcomponent =  request.GET.get('by_subcomponent')
    start_day =  request.GET.get('startdate')
    end_day =  request.GET.get('end')
 
    form = RepositoryIndicatorForm()
    repositoryIndicators = RepositoryIndicator.objects.all().order_by("-publish_date")
    
    
    if request.method=='GET':
        st=request.GET.get('indicator_code')
        if st!=None:
            repositoryIndicators= RepositoryIndicator.objects.filter(indicator_code__icontains=st).order_by("-publish_date")
            
        if by_subcomponent!=None:
            repositoryIndicators= RepositoryIndicator.objects.filter(Q(indicator__subcomponent__id__icontains=by_subcomponent)& Q(publish_date__gte=start_day)& Q(publish_date__lte=end_day)).order_by("-publish_date")
    
    context['repositoryIndicators'] = repositoryIndicators
    #context['paginator']=paginator
    context['title'] ='home'
    if request.method == 'POST':
        if 'save' in request.POST:
            pk = request.POST.get('save')
            if not pk:
                form=RepositoryIndicatorForm(request.POST)
            else:
                repositoryIndicator = RepositoryIndicator.objects.get(id = pk)
                form = RepositoryIndicatorForm(request.POST ,instance = repositoryIndicator) 
            form.save()
            form = RepositoryIndicatorForm()
            messages.success(request, 'Indicator value saved or updated successfully')
        elif 'delete' in request.POST:
            if request.user.groups.filter(name='Manager and delete').exists(): ## Condition security
                pk = request.POST.get('delete')
                repositoryIndicator = RepositoryIndicator.objects.get(id = pk)
                repositoryIndicator.delete()
                messages.success(request, 'Indicator value deleted successfully')
            else:
                return render(request, 'pages/examples/403.html')
    
        elif 'edit' in request.POST:
            pk = request.POST.get('edit')
            repositoryIndicator = RepositoryIndicator.objects.get(id = pk)
            form = RepositoryIndicatorForm(instance = repositoryIndicator)
          
    context['form'] = form
    context['form_select_repository'] = form_select_repository 
    context['by_subcomponent'] = by_subcomponent
    context['end_day'] = start_day
    context['end_day'] = end_day

    return render(request, 'mytvddata/pages/repository/repository_dataset.html', context)

# Method for list all Survey dataset submitted
#@permission_required('myuhp.view_surveydataset',raise_exception=True)

def repository_report(request):
    subComponents = Subcomponent.objects.all().order_by("component") 
    if request.method=='GET':
        st=request.GET.get('component')
        if st!=None:
            subComponents = Subcomponent.objects.filter(component__icontains=st)
    
    page = request.GET.get('page')
    num_of_items = 10
    # paginator of country 
    paginator = Paginator(subComponents, num_of_items)
       
    try:
        subComponents = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        subComponents = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        subComponents = paginator.page(page)
            
    data={'subComponents': subComponents, 'paginator':paginator}
    
    return render(request, 'mytvddata/pages/repository/repository_report.html', data)

# Method for view single Survey DataSet page

def single_repository_page(request, pk):
    try:
        subComponent=Subcomponent.objects.get(id=pk) 
        subComponents =Subcomponent.objects.filter(Q(id__icontains=pk) & Q(subcomponent_indicator__indicator_repository__indicator_code__isnull=False)).values(
        'subcomponent_name',
        'subcomponent_indicator__indicator_repository__country__name',
        'subcomponent_indicator__indicator_repository__spatial_dim',
        'subcomponent_indicator__indicator_repository__indicator__indicator_name',
        'subcomponent_indicator__indicator_repository__indicator_code',
        'subcomponent_indicator__indicator_repository__dim1_type',
        'subcomponent_indicator__indicator_repository__dim1',
        'subcomponent_indicator__indicator_repository__dim2_type',
        'subcomponent_indicator__indicator_repository__dim2',
        'subcomponent_indicator__indicator_repository__dim3_type',
        'subcomponent_indicator__indicator_repository__dim3',
        'subcomponent_indicator__indicator_repository__time_dim',
        'subcomponent_indicator__indicator_repository__alpha_value',
        'subcomponent_indicator__indicator_repository__numeric_value',
        'subcomponent_indicator__indicator_repository__publish_date'
        
    ).order_by('-subcomponent_indicator__indicator_repository__publish_date')
        
        
        page = request.GET.get('page')
        num_of_items = 10
        # paginator of country 
        paginator = Paginator(subComponents, num_of_items)
        
        try:
            subComponents = paginator.page(page)
        except PageNotAnInteger:
            page = 1
            subComponents = paginator.page(page)
        except EmptyPage:
            page = paginator.num_pages
            subComponents = paginator.page(page)
        
        data={'subComponent':subComponent, 'subComponents':subComponents, 'paginator':paginator}   
    except:
        data={'message':'The repository DataSet you requested doesn t\' exist or isn t in the submitted results'}
    
    return render(request,'mytvddata/pages/repository/single_repository_page.html',data)

#Method to export REPOSITORY DATASET on Excel file
def export_repository_to_excel(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Saving_indicator_results.xlsx"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Datawarehouse"

    # Add headers
    headers = ["sub_component","country", "spatial_dim", "indicator_name", "indicator_code", "dim1_type", "dim1", "dim2_type", "dim2", "dim3_type","dim3","time_dim","alpha_value","numeric_value","publish_date"]
    ws.append(headers)

    # Add data from the model
    repositoryIndicators = RepositoryIndicator.objects.all()
    for s in repositoryIndicators:
        ws.append([s.indicator.subcomponent.subcomponent_name,s.country.name, s.spatial_dim, s.indicator.indicator_name,s.indicator.indicator_code, s.dim1_type, s.dim1, 
                   s.dim2_type,  s.dim2,  s.dim3_type, s.dim3, s.time_dim, s.alpha_value, s.numeric_value, s.publish_date])

    # Save the workbook to the HttpResponse
    wb.save(response)
    return response

 
 
# =========================
# EXPORT REPOSITORY PAGE (PRO)
# =========================
@login_required
def export_repository_page(request, pk):
 
    # ✅ Réponse HTTP (.xlsx moderne)
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=repository_data.xlsx'
 
    wb = Workbook()
    ws = wb.active
    ws.title = "Repository Data"
 
    # =========================
    # HEADERS
    # =========================
    headers = [
        "Subcomponent", "Country", "Spatial Dim",
        "Indicator Name", "Indicator Code",
        "Dim1 Type", "Dim1", "Dim2 Type", "Dim2",
        "Dim3 Type", "Dim3", "Time",
        "Alpha Value", "Numeric Value", "Publish Date"
    ]
 
    ws.append(headers)
 
    # =========================
    # QUERY OPTIMISÉE
    # =========================
    subcomponent = Subcomponent.objects.prefetch_related(
        'subcomponent_indicator__indicator_repository'
    ).get(pk=pk)
 
    # =========================
    # DATA
    # =========================
    for indicator in subcomponent.subcomponent_indicator.all():
        for repo_data in indicator.indicator_repository.all():
 
            ws.append([
                subcomponent.subcomponent_name,
                repo_data.country.name if repo_data.country else '',
                repo_data.spatial_dim,
                indicator.indicator_name,
                repo_data.indicator_code,
                repo_data.dim1_type,
                repo_data.dim1,
                repo_data.dim2_type,
                repo_data.dim2,
                repo_data.dim3_type,
                repo_data.dim3,
                repo_data.time_dim,
                repo_data.alpha_value,
                repo_data.numeric_value,
                repo_data.publish_date
            ])
 
    # =========================
    # AUTO WIDTH (UX)
    # =========================
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
 
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
 
        ws.column_dimensions[col_letter].width = max_length + 2
 
    # ✅ SAVE
    wb.save(response)
 
    return response
 


## METHOD EXPORT THE FILTERING DATA UNIT IN EXCEL FORMAT
def export_to_excel_repositoring(request,by_subcomponent,end_day):
    #startdate
    """
    Une vue Django qui prend l'objet request et trois autres paramètres.
    """
 #Q(subcomponent_indicator__indicator_repository__publish_date__gte=start_day)&
       
    sub_component=Subcomponent.objects.get(id=by_subcomponent)
    data_component = Subcomponent.objects.all().filter(Q(id__icontains=by_subcomponent)&  Q(subcomponent_indicator__indicator_repository__publish_date__lte=end_day)).values(
        'subcomponent_name',
        'subcomponent_indicator__indicator_repository__country__name',
        'subcomponent_indicator__indicator_repository__spatial_dim',
        'subcomponent_indicator__indicator_repository__indicator__indicator_name',
        'subcomponent_indicator__indicator_repository__indicator_code',
        'subcomponent_indicator__indicator_repository__dim1_type',
        'subcomponent_indicator__indicator_repository__dim1',
        'subcomponent_indicator__indicator_repository__dim2_type',
        'subcomponent_indicator__indicator_repository__dim2',
        'subcomponent_indicator__indicator_repository__dim3_type',
        'subcomponent_indicator__indicator_repository__dim3',
        'subcomponent_indicator__indicator_repository__time_dim',
        'subcomponent_indicator__indicator_repository__alpha_value',
        'subcomponent_indicator__indicator_repository__numeric_value',
        'subcomponent_indicator__indicator_repository__publish_date'
        ).order_by('-subcomponent_indicator__indicator_repository__publish_date')
    
   # number_register =data_survey.annotate(total_sample=Count('project_surveyData__level_1')).values('total_sample')
    # Create fill
    
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="'+'Saving_dataWarehouse_'+sub_component.subcomponent_name+"_"+end_day+'.xlsx"'  
    workbook=Workbook()
    worksheet = workbook.active
    
    # Edit page Setup
    worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
    worksheet.page_setup.paperSize = worksheet.PAPERSIZE_TABLOID
    worksheet.page_setup.fitToHeight = 0
    worksheet.page_setup.fitToWidth = 1
    
    worksheet.merge_cells('A1:F1')
    worksheet.merge_cells('A2:F2')
    worksheet.freeze_panes = "F4"
    worksheet.auto_filter.ref = "A3:J5000"
    worksheet["H1"] = "TOTAL REGISTER"
    worksheet["I1"] = '=COUNTA(C4:C5000)'
    worksheet["I1"].font = Font(bold=True,color="0000ff", size=14)
    worksheet["I1"].alignment = Alignment(horizontal ="center", vertical="center")
 #   worksheet["F2"] = "NEW RISK"
 #   worksheet["G2"] = '=COUNTIF(G4:G200, "New")'
 #   worksheet["G2"].font = Font(bold=True, color="00FF0000", size=14, italic=True)
 #   worksheet["G2"].alignment = Alignment(horizontal ="center", vertical="center")
 #   worksheet["H2"] = "EXISTING RISK"
 #   worksheet["I2"] = '=COUNTIF(G4:G200, "Existing")'
 #   worksheet["I2"].font = Font(bold=True, color="00993300", size=14, italic=True)
 #   worksheet["I2"].alignment = Alignment(horizontal ="center", vertical="center")
 #   worksheet["J2"] = "CLOSED"
 #   worksheet["K2"] = '=COUNTIF(G4:G200, "closed")'
 #   worksheet["K2"].font = Font(bold=True, color="ff00ff", size=14, italic=True)
 #   worksheet["K2"].alignment = Alignment(horizontal ="center", vertical="center")
    
 #   worksheet["L2"] = "MODERATE"
 #   worksheet["M2"] = '=COUNTIF(J4:J200, "Moderate")'
 #   worksheet["M2"].font = Font(bold=True, color="0000CCFF", size=14, italic=True)
 #   worksheet["M2"].alignment = Alignment(horizontal ="center", vertical="center")
    
 #   worksheet["N2"] = "CRITICAL"
 #   worksheet["O2"] = '=COUNTIF(J4:J200, "Critical")'
 #   worksheet["O2"].font = Font(bold=True, color="00993366", size=14, italic=True)
 #   worksheet["O2"].alignment = Alignment(horizontal ="center", vertical="center")
    
 #   worksheet["P2"] = "VERY CRITICAL"
 ##   worksheet["Q2"] = '=COUNTIF(J4:J200, "Very_critical")'
 #   worksheet["Q2"].font = Font(bold=True, color="00FF6600", size=14, italic=True)
 #   worksheet["Q2"].alignment = Alignment(horizontal ="center", vertical="center")
    
 #  worksheet["R1"] = "MONITORING (%)"
 #   worksheet["S1"] ='=AVERAGE(R4:R20)'
 #   worksheet["S1"].font = Font(bold=True, color="0000FF", size=14, italic=True)
 #   worksheet["S1"].alignment = Alignment(horizontal ="center", vertical="center")
    #worksheet.cell['S1'].number_format = '0.00%'
 #   worksheet.conditional_formatting.add("S1", color_scale_rule)
    
   
  
  # Again, let's add this gradient to the star ratings, column "H"
  #  worksheet.conditional_formatting.add("R4:R20", color_scale_rule)
   
    first_cell = worksheet['A1']
    first_cell.value ="Repository: " +""+ sub_component.component.component_name +" and End date: "+end_day
    first_cell.fill = PatternFill("solid",fgColor="246ba1")
    first_cell.font =  Font(bold=True, color="F7F6FA", size=16)
    first_cell.alignment = Alignment(horizontal ="center", vertical="center")
     
    second_cell = worksheet['A2']
    second_cell.value = sub_component.subcomponent_name
    second_cell.font = Font(bold=True, color="246ba1", size=16)
    second_cell.alignment = Alignment(horizontal ="center", vertical="center")
    
   
        
    worksheet.title='sub_component.component.component_name' + "" +"_"+end_day
     # Define the titles for columns
    columns = [
               "sub_component","country", "spatial_dim", "indicator_name",
               "indicator_code", "dim1_type", "dim1", "dim2_type", "dim2", 
               "dim3_type","dim3","time_dim","alpha_value","numeric_value",
               "publish_date"
                ]
 
    row_num = 3
    
    # Assign the titles for each cell of the header
    for col_num, column_title in enumerate(columns, 1):
        cell = worksheet.cell(row=row_num, column=col_num)
        cell.value= column_title
        cell.fill = PatternFill("solid", fgColor="50c878")
        cell.font = Font(bold=True, color="F7F6FA")
        cell.alignment = Alignment(horizontal ="center", vertical="center",wrap_text=False)
        third_cell=worksheet['F3']
        third_cell.alignment =Alignment(horizontal="center", vertical="center",wrap_text=False)
        
    for subComponent in data_component:
        row_num +=1    
        #Define the data for each cell in the row
        row=[subComponent['subcomponent_name'],
             subComponent['subcomponent_indicator__indicator_repository__country__name'],
             subComponent['subcomponent_indicator__indicator_repository__spatial_dim'],
             subComponent['subcomponent_indicator__indicator_repository__indicator__indicator_name'],
             subComponent['subcomponent_indicator__indicator_repository__indicator_code'],
             subComponent['subcomponent_indicator__indicator_repository__dim1_type'],
             subComponent['subcomponent_indicator__indicator_repository__dim1'],
             subComponent['subcomponent_indicator__indicator_repository__dim2_type'],
             subComponent['subcomponent_indicator__indicator_repository__dim2'],
             subComponent['subcomponent_indicator__indicator_repository__dim3_type'],
             subComponent['subcomponent_indicator__indicator_repository__dim3'],
             subComponent['subcomponent_indicator__indicator_repository__time_dim'],
             subComponent['subcomponent_indicator__indicator_repository__alpha_value'],
             subComponent['subcomponent_indicator__indicator_repository__numeric_value'],
             subComponent['subcomponent_indicator__indicator_repository__publish_date']
             ]
                   
          #Assign the data for each cell in the row
        for col_num, cell_value in enumerate(row, 1):
            cell = worksheet.cell(row=row_num, column=col_num)
            cell.value = cell_value
            cell.alignment =Alignment(vertical="center",wrap_text=False)
    #          if isinstance(cell_value, decimal.Decimal):
    #              cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
    workbook.save(response)
    return response

    
def upload_success_repo(request):
    return render(request, 'mytvddata/pages/repository/upload_success_repo.html')


def dataset_list_reposi(request):
    repository_indicators = RepositoryIndicator.objects.all()
    return render(request, 'mytvddata/pages/repository/dataset_list_reposi.html', {'repository_indicators': repository_indicators})


def display_api_data(request):
    
    context={}
    form_select_api= SelectAPIForm(request.GET or None)
    by_indicator =  request.GET.get('by_indicator')
    start_day =  request.GET.get('startdate')
    end_day =  request.GET.get('end')

    form = StoreAPIForm()
    storeAPIs = StoreAPI.objects.all().order_by("-publish_date")
    
    if request.method=='GET':
        st=request.GET.get('indicator_code')
        if st!=None:
            storeAPIs= StoreAPI.objects.filter(indicator_code__icontains=st).order_by("-publish_date")
        
        if by_indicator!=None:
            storeAPIs= StoreAPI.objects.filter(Q(indicator_code__icontains=by_indicator)& Q(publish_date__gte=start_day)& Q(publish_date__lte=end_day)).order_by("-publish_date")
    
    page = request.GET.get('page')
    num_of_items = 100
    # paginator of country 
    paginator = Paginator(storeAPIs, num_of_items)
       
    try:
        storeAPIs = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        storeAPIs = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        storeAPIs = paginator.page(page)
            
    context['storeAPIs'] = storeAPIs
    context['paginator']=paginator
    context['title'] ='home'
    if request.method == 'POST':
        if 'save' in request.POST:
            pk = request.POST.get('save')
            if not pk:
                form=StoreAPIForm(request.POST)
            else:
                storeAPI = StoreAPI.objects.get(id = pk)
                form = StoreAPIForm(request.POST ,instance = storeAPI) 
            form.save()
            form = StoreAPIForm()
            messages.success(request, 'API data saved or updated successfully')
        elif 'delete' in request.POST:
          #  if request.user.groups.filter(name='Manager and delete').exists(): ## Condition security
                pk = request.POST.get('delete')
                storeAPI = StoreAPI.objects.get(id = pk)
                storeAPI.delete()
                messages.success(request, 'API data deleted successfully')
           # else:
          #      return render(request, 'pages/examples/403.html')
    
        elif 'edit' in request.POST:
            pk = request.POST.get('edit')
            storeAPI = StoreAPI.objects.get(id = pk)
            form = StoreAPIForm(instance = storeAPI)
          
    context['form'] = form
    context['form_select_api'] = form_select_api 
    context['by_indicator'] = by_indicator
    context['end_day'] = start_day
    context['end_day'] = end_day
    
   # data_from_db = StoreAPI.objects.all().order_by('-created_at')
    return render(request, 'mytvddata/epidata/api_dataset.html', context)






@csrf_exempt
def load_api_json(request):
   
    api_url={}
    api_indicator={}
    api_data_period={}
    details={}  
    context={}  
    payload={}
   
   
    data_from_db=[]
   
    form_select_url= SelectUrlForm(request.GET or None)
    by_indicator =  request.GET.get('by_indicator')
    start_day =  request.GET.get('startdate')
    end_day =  request.GET.get('end')
   
    ##https://ghoapi.azureedge.net/api/MALARIA_INDIG?$filter=ParentLocationCode%20eq%20%27AFR%27and%20date(TimeDimensionBegin)%20ge%202023-01-01%20and%20date(TimeDimensionBegin)%20lt%202026-01-01
   
   
    headers = {"Content-Type": "application/json", "Accept-Encoding": "deflate"}
   
    base_url="https://ghoapi.azureedge.net/api/"
 
    if request.method=='GET':
       
        url=f"{base_url}/{by_indicator}?$filter=ParentLocationCode eq 'AFR'and date(TimeDimensionBegin) ge {start_day} and date(TimeDimensionBegin) lt {end_day}"
        if by_indicator!=None:
       # st= request.GET.get("api_url") # 1. Appeler l'API JSON
      #  st= request.GET.get("api_url") # 1. Appeler l'API JSON
      #  if st!=None:
            #response = requests.get(url,header, verify=False)
            response = requests.get(url, headers=headers, verify=False)
            response.raise_for_status()  # Lève une exception pour les codes d'état d'erreur
 
            if response.status_code==200:  # Vérifier que la requête est OK
               payload=json.loads(response.text)
               details = payload.get('value')
               
               for item in details:
                # Insert into MySQL via Django ORM
                StoreAPI.objects.update_or_create(
                    api_id=item['Id'],  # unique field
                    defaults={
                        "indicator_code": item["IndicatorCode"],
                        "country_code": item["SpatialDim"],
                        "dim1_type": item["Dim1Type"],
                        "dim1": item["Dim1"],
                        "time_dim": item["TimeDim"],
                        "dim2_type": item["Dim2Type"],
                        "dim2": item["Dim2"],
                        "dim3_type": item["Dim3Type"],
                        "dim3": item["Dim3"],
                        "alpha_value": item["Value"],
                        "numeric_value": item["NumericValue"],
                        "publish_date": item["Date"]
                    }
                )
                messages.success(request, 'API data from GHO saved or updated successfully')
            else:
               # print("Failed to fetch data:")
                messages.warning(request, 'Failed to fetch data:')
               
        data_from_db = StoreAPI.objects.all().order_by('-created_at')
        context['api_data'] = data_from_db
        context['form_select_url'] = form_select_url
    return render(request, 'mytvddata/cover/api_data_list.html', context)
       
 
@csrf_exempt
def loads_api(request): # GOOD Method to store data from API to MySQL database USE WITH POSTMAN PLATFORM

    payload=json.loads(request.body)
    data = payload.get('value')
        
        #    url = "https://api.example.com/users"  # Replace with your API
        #    response = requests.get(url)
    if request.method == 'POST':
            
     #   if data.status_code == 200:
     #       data = response.json()  # JSON → Python dict/list
            
            for item in data:
                # Insert into MySQL via Django ORM
                ApiData.objects.update_or_create(
                    api_id=item['Id'],  # unique field
                    defaults={
                        "IndicatorCode": item["IndicatorCode"],
                        "SpatialDimType": item["SpatialDimType"],
                        "SpatialDim": item["SpatialDim"],
                        "TimeDimType": item["TimeDimType"],
                        "ParentLocationCode": item["ParentLocationCode"],
                        "ParentLocation": item["ParentLocation"],
                        "Dim1Type": item["Dim1Type"],
                        "Dim1": item["Dim1"],
                        "TimeDim": item["TimeDim"],
                        "Dim2Type": item["Dim2Type"],
                        "Dim2": item["Dim2"],
                        "Dim3Type": item["Dim3Type"],
                        "Dim3": item["Dim3"],
                        "DataSourceDim": item["DataSourceDim"],
                        "Value": item["Value"],
                        "NumericValue": item["NumericValue"],
                        "Low": item["Low"],
                        "High": item["High"],
                        "Comments": item["Comments"],
                        "Date": item["Date"],
                        "TimeDimensionValue": item["TimeDimensionValue"],
                        "TimeDimensionBegin": item["TimeDimensionBegin"],
                        "TimeDimensionEnd": item["TimeDimensionEnd"]
                    }
                )
    else:
        print("Failed to fetch data:")
        
    return HttpResponse("Data successfully fetched and stored!")    


# Optional: View to display the fetched GHO from your database
def fetch_api_stored(request):
    response={}  
    payload={}
    header={"Content-Type":"applica/json",
            "Accept-Encoding":"Deflate"}
    
    if request.method=='GET':
        st=request.GET.get("api_url")  # 1. Appeler l'API JSON
        if st!=None:
            response = requests.get(st,header, verify=False)
            response.raise_for_status()  # Lève une exception pour les codes d'état d'erreur

            if response.status_code==200:  # Vérifier que la requête est OK
               payload=json.loads(response.text)
             #  details = payload.get('value')
               return JsonResponse({
                     'payload':payload
               }, safe=False)
                
            else:
                messages.info(request, "Erreur lors de l'appel AP:", response.status_code)
    return JsonResponse({
                     'response':response
               }, safe=False)
                
  #  return HttpResponse("Données de l'API chargées avec succès dans la base de données!")
 
        
# Method for upload pdf document or report in pdf format
#@lockdown()
#@permission_required('myuhp.add_reportsave',raise_exception=True)
def report_upload(request):
    form =None
    if request.method == "POST":
        form = ReportSaveForm(request.POST, request.FILES)
       # print(form)
        if form.is_valid():
            form.save()
            messages.success(request, 'pdf file added successfully')
            # Get the current instance object to display in the template
            img_report_obj = form.instance
        return render(request, 'mytvddata/pages/report/report_upload.html', {'form': form, 'img_report_obj': img_report_obj})
    else:
        form = ReportSaveForm()
        img_report_obj_all = ReportSave.objects.all()
     
    return render(request,'mytvddata/pages/report/report_upload.html', {'form': form, 'img_report_obj_all':img_report_obj_all})


# Method for views all reports upload in the system
def views_report(request):
    reportSaves = ReportSave.objects.all().order_by("title_rep") 
    if request.method=='GET':
        st=request.GET.get('title_rep')
        if st!=None:
            reportSaves= ReportSave.objects.filter(title_rep__icontains=st)
    
    page = request.GET.get('page')
    num_of_items = 10
    # paginator of country 
    paginator = Paginator(reportSaves, num_of_items)
       
    try:
        reportSaves = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        reportSaves = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        reportSaves = paginator.page(page)
            
    data={'reportSaves': reportSaves, 'paginator':paginator}
    
    return render(request, 'mytvddata/pages/report/report_list.html', data)

# Method to put  into the table all reports upload in the system
#@login_required
#@permission_required('myuhp.view_reportsave',raise_exception=True)
def index_report(request):
    reportSaves = ReportSave.objects.all().order_by("title_rep") 
    if request.method=='GET':
        st=request.GET.get('title_rep')
        if st!=None:
            reportSaves= ReportSave.objects.filter(title_rep__icontains=st)
    
    page = request.GET.get('page')
    num_of_items = 10
    # paginator of country 
    paginator = Paginator(reportSaves, num_of_items)
       
    try:
        reportSaves = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        reportSaves = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        reportSaves = paginator.page(page)
            
    data={'reportSaves': reportSaves, 'paginator':paginator}
    
    return render(request, 'mytvddata/pages/report/report_index.html', data)

# Edit report thad added in the system
#@login_required
#@permission_required('myuhp.add_reportsave',raise_exception=True)
def editReport(request, pk):
    reportSaves = ReportSave.objects.get(id=pk)
    form = ReportSaveForm(instance=reportSaves)
    
    if request.method == "POST":
        form = ReportSaveForm(request.POST, request.FILES, instance=reportSaves)
        if form.is_valid():
            form.save()
            messages.success(request, "Report updated Successfully")
            return redirect('index_report')

    messages.warning(request, "ERROR FORM INVALID")
    return render(request, 'mytvddata/pages/report/report_edit.html', {'form': form, 'reportSaves':reportSaves})


# Method for upload document or image
#@lockdown()
#@permission_required('myuhp.add_docsave',raise_exception=True)
def docSave_upload(request):
    form =None
    if request.method == "POST":
        form = DocSaveForm(request.POST, request.FILES)
       # print(form)
        if form.is_valid():
            form.save()
            # Get the current instance object to display in the template
            img_obj = form.instance
        return render(request, 'mytvddata/pages/report/doc_upload.html', {'form': form, 'img_obj': img_obj})
    else:
        form = DocSaveForm()
        img_obj_all = DocSave.objects.all()
     
    return render(request,'mytvddata/pages/report/doc_upload.html', {'form': form, 'img_obj_all':img_obj_all})


## METHOD EXPORT THE FILTERING DATA UNIT IN EXCEL FORMAT
def export_to_excel_API(request,by_indicator,end_day):
    #startdate
    """
    Une vue Django qui prend l'objet request et trois autres paramètres.
    """
 #Q(subcomponent_indicator__indicator_repository__publish_date__gte=start_day)&
       
   # sub_component=StoreAPI.objects.get(indicator_code=by_indicator)
    data_component = StoreAPI.objects.all().filter(Q(indicator_code__icontains=by_indicator)&  Q(publish_date__lte=end_day)).values(
        'api_id',
        'country_code',
        'indicator_code',
        'dim1_type',
        'dim1',
        'dim2_type',
        'dim2',
        'dim3_type',
        'dim3',
        'time_dim',
        'alpha_value',
        'numeric_value',
        'publish_date'
        ).order_by('-publish_date')
    
   # number_register =data_survey.annotate(total_sample=Count('project_surveyData__level_1')).values('total_sample')
    # Create fill
    
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="'+'Saving_api_'+by_indicator+"_"+end_day+'.xlsx"'  
    workbook=Workbook()
    worksheet = workbook.active
    
    # Edit page Setup
    worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
    worksheet.page_setup.paperSize = worksheet.PAPERSIZE_TABLOID
    worksheet.page_setup.fitToHeight = 0
    worksheet.page_setup.fitToWidth = 1
    
    worksheet.merge_cells('A1:F1')
    worksheet.merge_cells('A2:F2')
    worksheet.freeze_panes = "F4"
    worksheet.auto_filter.ref = "A3:J5000"
    worksheet["H1"] = "TOTAL REGISTER"
    worksheet["I1"] = '=COUNTA(C4:C5000)'
    worksheet["I1"].font = Font(bold=True,color="0000ff", size=14)
    worksheet["I1"].alignment = Alignment(horizontal ="center", vertical="center")
 
   
    first_cell = worksheet['A1']
    first_cell.value ="API: " +""+ by_indicator +" and End date: "+end_day
    first_cell.fill = PatternFill("solid",fgColor="246ba1")
    first_cell.font =  Font(bold=True, color="F7F6FA", size=16)
    first_cell.alignment = Alignment(horizontal ="center", vertical="center")
     
   
        
    worksheet.title='by_indicator' + "" +"_"+end_day
     # Define the titles for columns
    columns = [
               "api_id", "country_code", "indicator_code","dim1_type",
                "dim1","dim2_type","dim2","dim3_type","dim3","time_dim",
                "alpha_value","numeric_value","publish_date"
                ]
 
    row_num = 3
    
    # Assign the titles for each cell of the header
    for col_num, column_title in enumerate(columns, 1):
        cell = worksheet.cell(row=row_num, column=col_num)
        cell.value= column_title
        cell.fill = PatternFill("solid", fgColor="50c878")
        cell.font = Font(bold=True, color="F7F6FA")
        cell.alignment = Alignment(horizontal ="center", vertical="center",wrap_text=False)
        third_cell=worksheet['F3']
        third_cell.alignment =Alignment(horizontal="center", vertical="center",wrap_text=False)
        
    for subComponent in data_component:
        row_num +=1    
        #Define the data for each cell in the row
        row=[subComponent['api_id'],
             subComponent['country_code'],
             subComponent['indicator_code'],
             subComponent['dim1_type'],
             subComponent['dim1'],
             subComponent['dim2_type'],
             subComponent['dim2'],
             subComponent['dim3_type'],
             subComponent['dim3'],
             subComponent['time_dim'],
             subComponent['alpha_value'],
             subComponent['numeric_value'],
             subComponent['publish_date']
             ]
                   
          #Assign the data for each cell in the row
        for col_num, cell_value in enumerate(row, 1):
            cell = worksheet.cell(row=row_num, column=col_num)
            cell.value = cell_value
            cell.alignment =Alignment(vertical="center",wrap_text=False)
    #          if isinstance(cell_value, decimal.Decimal):
    #              cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
    workbook.save(response)
    return response


# def component(request):
#    form=ComponentForm()
#    context={'form':form}
#    return render(request, 'mytvddata/pages/component/component.html', context)
    
#def subcomponents(request):
#    form=ComponentForm(request.GET)
#    return HttpResponse(form['subcomponents'])
    
# Method for CRUD Component
from django.shortcuts import redirect

def component_add(request):
    context = {}
    form = ComponentsForm()
    components = Component.objects.all()

    if request.method == 'GET':
        st = request.GET.get('component_name')
        if st:
            components = Component.objects.filter(component_name__icontains=st)

    page = request.GET.get('page')
    num_of_items = 10
    paginator = Paginator(components, num_of_items)

    try:
        components = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        components = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        components = paginator.page(page)

    if request.method == 'POST':
        if 'save' in request.POST:
            pk = request.POST.get('save')
            if not pk:
                form = ComponentsForm(request.POST)
            else:
                component = Component.objects.get(id=pk)
                form = ComponentsForm(request.POST, instance=component)
            if form.is_valid():
                form.save()
                messages.success(request, 'Component saved or updated successfully')
            return redirect('component_add')  # 🔑 redirection après POST

        elif 'delete' in request.POST:
            pk = request.POST.get('delete')
            component = Component.objects.get(id=pk)
            component.delete()
            messages.success(request, 'Component deleted successfully')
            return redirect('component_add')  # 🔑 redirection après POST

        elif 'edit' in request.POST:
            pk = request.POST.get('edit')
            component = Component.objects.get(id=pk)
            form = ComponentsForm(instance=component)

    context['components'] = components
    context['paginator'] = paginator
    context['title'] = 'home'
    context['form'] = form

    return render(request, 'mytvddata/pages/component/component.html', context)


#Method to export ALL COMPONENT on Excel file
def export_component(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="Component.xlsx"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Component"

    # Add headers
    headers = ["id","component_name"]
    ws.append(headers)

    # Add data from the model
    components = Component.objects.all()
    for component in components:
        ws.append([component.pk, component.component_name])

    # Save the workbook to the HttpResponse
    wb.save(response)
    return response


# Method for CRUD Component
def subcomponent_add(request):
    context={}
    form = SubcomponentsForm()
    subcomponents = Subcomponent.objects.all()
    
    if request.method=='GET':
        st=request.GET.get('subcomponent_name')
        if st!=None:
            subcomponents= Subcomponent.objects.filter(subcomponent_name__icontains=st)
    
    page = request.GET.get('page')
    num_of_items = 100
    # paginator of country 
    paginator = Paginator(subcomponents, num_of_items)
       
    try:
        subcomponents = paginator.page(page)
    except PageNotAnInteger:
        page = 1
        subcomponents = paginator.page(page)
    except EmptyPage:
        page = paginator.num_pages
        subcomponents = paginator.page(page)
            
    context['subcomponents'] = subcomponents
  
    context['paginator']=paginator
    context['title'] ='home'
    if request.method == 'POST':
        if 'save' in request.POST:
            pk = request.POST.get('save')
            if not pk:
                form=SubcomponentsForm(request.POST)
            else:
                subcomponent = Subcomponent.objects.get(id = pk)
                form = SubcomponentsForm(request.POST ,instance = subcomponent) 
            if form.is_valid():
                form.save()
                messages.success(request, 'Disease saved or updated successfully')
            return redirect('subcomponent_add')  # 🔑 redirection après POST

        elif 'delete' in request.POST:
         #   if request.user.groups.filter(name='Manager and delete').exists(): ## Condition security
                pk = request.POST.get('delete')
                subcomponent = Subcomponent.objects.get(id = pk)
                subcomponent.delete()
                messages.success(request, 'Sub Component deleted successfully')
                return redirect('subcomponent_add')  # 🔑 redirection après POST
        #    else:
        #        return render(request, 'pages/examples/403.html')
    
        elif 'edit' in request.POST:
            pk = request.POST.get('edit')
            subcomponent = Subcomponent.objects.get(id = pk)
            form = SubcomponentsForm(instance = subcomponent)
          
    context['form'] = form
    context['paginator'] = paginator
    context['title'] = 'home'
  
    return render(request, 'mytvddata/pages/component/subcomponent.html', context)


#Method to export ALL SUB-COMPONENT on Excel file
def export_subcomponent(request):
    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="SubComponent.xlsx"'

    wb = Workbook()
    ws = wb.active
    ws.title = "SubComponent"

    # Add headers
    headers = ["id","component","subcomponent_name"]
    ws.append(headers)

    # Add data from the model
    subcomponents = Subcomponent.objects.all()
    for subcomponent in subcomponents:
        ws.append([subcomponent.pk, subcomponent.component.component_name,subcomponent.subcomponent_name])

    # Save the workbook to the HttpResponse
    wb.save(response)
    return response

# CREATE THE JSON DATA FOR INDICATOR
@csrf_exempt
def json_indicator(request):
    data=list(Indicator.objects.values('indicator_code','indicator_name','indicator_target','indicator_metric','type_indicator','ref_data',
                                       'indicator_unit','indicator_source','category_indicator','forecasting_indicator','performance_indicator','subcomponent__subcomponent_name','subcomponent__component__component_name'))
    return JsonResponse(data, safe=False)


# CREATE THE JSON DATA FOR STOREAPI
@csrf_exempt
def json_storeAPI(request):
    data_storeAPI=list(StoreAPI.objects.values('indicator_code','country_code','time_dim','dim1_type','dim1',
                                               'dim2_type','dim2','dim3_type','dim3','alpha_value','numeric_value'))
    return JsonResponse(data_storeAPI, safe=False)


# CREATE THE JSON DATA FOR REPOSITORY
@csrf_exempt
def json_repository(request):
    data_repository=list(RepositoryIndicator.objects.values('indicator__subcomponent__component__component_name','indicator__subcomponent__subcomponent_name','country__name','spatial_dim', 'indicator__indicator_name',
                                                            'indicator__indicator_code','dim1_type', 'dim1', 'dim2_type', 'dim2',  'dim3_type', 'dim3', 'time_dim', 
                                                            'alpha_value', 'numeric_value', 'publish_date'))  
    return JsonResponse(data_repository, safe=False)


# CREATE THE JSON DATA FOR COUNTRY
@csrf_exempt
def json_country(request):
    data_country=list(Country.objects.values('flag','cca3','name','official','capital','subregion',
                                             'area','population','languages','ref_data','data_source','country_class'))
    return JsonResponse(data_country, safe=False)


# CREATE THE JSON DATA FOR LOCATION GPS
@csrf_exempt
def json_location(request):
    data_location=list(LocationCountry.objects.values('iso3','name','latitude','longitude'))
    return JsonResponse(data_location, safe=False)

# CREATE THE JSON DATA FOR COMPONENT
@csrf_exempt
def json_component(request):
    data_component=list(Component.objects.values('id','component_name'))
    return JsonResponse(data_component, safe=False)


# CREATE THE JSON DATA FOR SUB COMPONENT
@csrf_exempt
def json_subcomponent(request):
    data_subcomponent=list(Subcomponent.objects.values('id','subcomponent_name','component__component_name'))
    return JsonResponse(data_subcomponent, safe=False)

# CREATE THE JSON DATA FOR SURVEY PROJECT
@csrf_exempt
def json_surveyproject(request):
    data_surveyproject=list(SurveyProject.objects.values('id','responsible','title_surv',
                                                         'target_population',
                                                         'start_date',
                                                         'end_date',
                                                         'location_survey',
                                                         'date_creation'))
    return JsonResponse(data_surveyproject, safe=False)

# CREATE THE JSON DATA FOR SURVEY DATASET
@csrf_exempt
def json_surveydataset(request):
    data_survey=list(SurveyDataset.objects.values('id','surveyProject__responsible',
                                                  'surveyProject__title_surv',
                                                  'surveyProject__start_date',
                                                  'surveyProject__end_date',
                                                  'surveyProject__location_survey',
                                                  'quest_code','question','response_text','response_num',
                                                  'level_1','level_2'))
    return JsonResponse(data_survey, safe=False)


# CREATE THE JSON DATA FOR USER
@csrf_exempt
def json_user(request):
    data_user=list(User.objects.values('username', 'email'))
    return JsonResponse(data_user, safe=False)


 #DISPLAY PAGE OF ERROR 500
def custom_500(request):
    return render(request, 'mytvddata/cover/500.html', status=500)

 #DISPLAY PAGE OF ERROR 404
def handel404(request, exception):
    return render(request, 'mytvddata/cover/404.html')

# DISPLAY PAGE OF ERROR 403
def handel403(request, exception):
    return render(request, 'mytvddata/cover/403.html')


def select_api_fields(request):
    available_fields = []
    sample_data = []
    details={}  
    context={}  
    payload={}
    db_fields = [f.name for f in ApiExtract._meta.get_fields() if f.concrete and not f.auto_created]

    form = SelectUrlForms(request.POST or None)


    if request.method == "POST" and form.is_valid():
        api_url = form.cleaned_data["api_url"]
        action = request.POST.get("action")

        if action == "import":
            message = import_api_data(api_url)
            messages.success(request, message)

        try:
            response = requests.get(api_url, timeout=5)
            response.raise_for_status()
            payload = response.json()
            details = payload.get("value", [])

            if details:
                # Extraire les champs disponibles
                first_item = details[0]
                available_fields = [
                    {"id": key, "field_name": key, "enabled": True}
                    for key in first_item.keys()
                ]

                # Sauvegarde en base (optionnel)
                for key in first_item.keys():
                    ApiFieldConfig.objects.get_or_create(field_name=key)

                # Prendre les 3 premières lignes
                sample_data = details[:3]

                 # Sauvegarde du mapping choisi
                for field in available_fields:
                    mapping_value = request.POST.get(f"mapping_{field['id']}")
                    if mapping_value:
                        ApiFieldMapping.objects.update_or_create(
                            api_field=field["field_name"],
                            defaults={"db_field": mapping_value}
                        )


                # Import des données si demandé
                if action == "import":
                    message = import_api_data(api_url)
                    messages.success(request, f"{message} Cliquez ci-dessous pour voir les données importées.")

        except Exception as e:
            available_fields = [{"id": "error", "field_name": f"Erreur: {e}", "enabled": False}]
            messages.error(request, f"Erreur lors de la récupération : {e}")

    context = {
        "form": form,
        "available_fields": available_fields,
        "sample_data": sample_data,
        "db_fields": db_fields,
    }

    return render(request, "mytvddata/cover/select_fields.html", context)



def import_api_data(api_url):
    response = requests.get(api_url, timeout=10)
    response.raise_for_status()
    payload = response.json()
    details = payload.get("value", [])

    if not details:
        return "Aucune donnée trouvée."

    # Charger le mapping API → DB
    mappings = {m.api_field: m.db_field for m in ApiFieldMapping.objects.all()}

    # Parcourir les lignes de l’API
    for item in details:
        record_data = {}
        for api_field, db_field in mappings.items():
            if db_field and api_field in item:
                record_data[db_field] = item[api_field]

        # Créer ou mettre à jour l’enregistrement
        ApiExtract.objects.create(**record_data)

    return f"{len(details)} lignes importées avec succès."



def view_imported_data(request):
    records = ApiExtract.objects.all()
    return render(request, "mytvddata/cover/imported_data.html", {
        "records": records
    })


def load_api_other(request):
    api_url = request.session.get("api_url")  # récupère l’URL saisie par l’utilisateur
    if not api_url:
        messages.error(request, "Veuillez d’abord saisir une URL API.")
        return redirect("select_api_fields")

    headers = {"Content-Type": "application/json"}
    allowed_fields = list(ApiFieldConfig.objects.filter(enabled=True).values_list("field_name", flat=True))

    try:
        response = requests.get(api_url, headers=headers, verify=False)
        response.raise_for_status()
        payload = response.json().get("value", [])

        for item in payload:
            filtered_data = {f: item.get(f) for f in allowed_fields}

            ApiExtract.objects.update_or_create(
                api_id=item.get("Id"),
                defaults={
                    "indicator_code": filtered_data.get("IndicatorCode"),
                    "country_code": filtered_data.get("SpatialDim"),
                    "parent_location": filtered_data.get("ParentLocation"),
                    "time_dim": filtered_data.get("TimeDim"),
                    "value": filtered_data.get("Value"),
                    "numeric_value": filtered_data.get("NumericValue"),
                    "date": filtered_data.get("Date"),
                }
            )
        messages.success(request, "API data saved successfully!")

    except Exception as e:
        messages.error(request, f"Error fetching API: {e}")

    data_from_db = ApiExtract.objects.all().order_by("-date")
    return render(request, "mytvddata/cover/api_data_other.html", {"api_data": data_from_db})

## FUnction for StaffMaping

### GOOD -------------------------# Liste principale avec formulaire vide pour le modal "create"
def staff_list(request):
    staff = StaffMember.objects.all().prefetch_related('diseases', 'country')

    # Filtres GET
    country_filter = request.GET.get("country")
    language_filter = request.GET.get("language")
    disease_filter = request.GET.get("disease")

    if country_filter:
        staff = staff.filter(country__name__icontains=country_filter)

    if language_filter:
        staff = staff.filter(language__icontains=language_filter)  # JSONField ou TextField

    if disease_filter:
        staff = staff.filter(diseases__subcomponent_name__icontains=disease_filter).distinct()


    # Distincts pour statistiques
    all_countries = [c.name for s in StaffMember.objects.all() for c in s.country.all()]
    distinct_countries = set(all_countries)

    all_languages = [l for s in StaffMember.objects.all() if s.language for l in s.language]
    distinct_languages = set(all_languages)

    all_diseases = [d.subcomponent_name for s in StaffMember.objects.all() for d in s.diseases.all()]
    distinct_diseases = set(all_diseases)

    form = StaffMemberForm()

    return render(
        request,
        "mytvddata/pages/staff/staff_list.html",
        {
            "staff": staff,
            "form": form,
            "countries": distinct_countries,
            "languages": distinct_languages,
            "diseases": distinct_diseases,
        }
    )

def staff_create(request):
    if request.method == "POST":
        form = StaffMemberForm(request.POST)
        if form.is_valid():
            staff = form.save()
            send_mail("Nouveau Staff ajouté", f"{staff.name} a été ajouté.", "admin@server.com", [staff.email], fail_silently=True)
            messages.success(request, "Staff ajouté avec succès.")
            return HttpResponse("OK")
        return render(request, "mytvddata/pages/staff/partial_staff_form.html", {"form": form})
    else:
        form = StaffMemberForm()
        return render(request, "mytvddata/pages/staff/partial_staff_form.html", {"form": form})

def staff_update(request, pk):
    staff = get_object_or_404(StaffMember, pk=pk)
    if request.method == "POST":
        form = StaffMemberForm(request.POST, instance=staff)
        if form.is_valid():
            staff = form.save()
            send_mail("Mise à jour Staff", f"{staff.name} a été mis à jour.", "admin@server.com", [staff.email], fail_silently=True)
            messages.success(request, "Staff mis à jour avec succès.")
            return HttpResponse("OK")
        return render(request, "mytvddata/pages/staff/partial_staff_form.html", {"form": form, "staff": staff})
    else:
        form = StaffMemberForm(instance=staff)
        return render(request, "mytvddata/pages/staff/partial_staff_form.html", {"form": form, "staff": staff})

def staff_delete(request, pk):
    staff = get_object_or_404(StaffMember, pk=pk)
    if request.method == "POST":
        staff.delete()
        messages.success(request, "Staff supprimé.")
    return redirect("staff_list")





def staff_export_xlsx(request):
    staff = StaffMember.objects.all().prefetch_related("diseases", "country")

    data = []
    for s in staff:
        data.append({
            "Name": s.name,
            "Email": s.email,
            "Countries": ", ".join([c.name for c in s.country.all()]),
            "Position": s.position,
            "Grade": s.grade,
            "Telephone": s.telephone,
            "Office Affiliation": s.office_affiliation,
            "Responsibility": s.responsibility,
            "Languages": ", ".join(s.language if s.language else []),
            "Diseases": ", ".join([d.subcomponent_name for d in s.diseases.all()]),
            "Level Geo": s.level_geo,
        })

    df = pd.DataFrame(data)

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = 'attachment; filename="staff.xlsx"'

    with pd.ExcelWriter(response, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Staff")

    return response







def staff_import_xlsx(request):
    if request.method == "POST" and request.FILES.get("file"):
        excel_file = request.FILES["file"]
        df = pd.read_excel(excel_file)

        for _, row in df.iterrows():
            staff, created = StaffMember.objects.update_or_create(
                email=row["Email"],
                defaults={
                    "name": row["Name"],
                    "position": row.get("Position", ""),
                    "grade": row.get("Grade", ""),
                    "telephone": row.get("Telephone", ""),
                    "office_affiliation": row.get("Office Affiliation", ""),
                    "responsibility": row.get("Responsibility", ""),
                    "level_geo": row.get("Level Geo", ""),
                }
            )

            # Pays (ManyToMany)
            if "Countries" in row and pd.notna(row["Countries"]):
                countries = [c.strip() for c in row["Countries"].split(",")]
                staff.country.set(Country.objects.filter(name__in=countries))

            # Langues (JSONField)
            if "Languages" in row and pd.notna(row["Languages"]):
                staff.language = [l.strip() for l in row["Languages"].split(",")]

            # Maladies (ManyToMany)
            if "Diseases" in row and pd.notna(row["Diseases"]):
                diseases = [d.strip() for d in row["Diseases"].split(",")]
                staff.diseases.set(Subcomponent.objects.filter(name__in=diseases))

            staff.save()

        messages.success(request, "Importation XLSX réussie.")
    return redirect("staff_list")





def staff_template_xlsx(request):
    # Exemple de colonnes attendues
    columns = [
        "Name", "Email", "Countries", "Position", "Grade",
        "Telephone", "Office Affiliation", "Responsibility",
        "Languages", "Diseases", "Level Geo"
    ]

    # Exemple de lignes vides ou pré-remplies
    df = pd.DataFrame(columns=columns)
    df.loc[0] = [
        "Jean Dupont", "jean.dupont@example.com", "France",
        "Analyste", "A1", "+33 123456789", "OMS Paris",
        "Responsable indicateurs", "Français, Anglais",
        "Diabète, Malaria", "Europe"
    ]
    df.loc[1] = [
        "Maria Sanchez", "maria.sanchez@example.com", "Espagne",
        "Coordinatrice", "B2", "+34 987654321", "OMS Madrid",
        "Coordination projets", "Espagnol",
        "Tuberculose", "Europe"
    ]

    # Réponse HTTP avec fichier Excel
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = 'attachment; filename="staff_template.xlsx"'

    with pd.ExcelWriter(response, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Template")

    return response

## DASHBOARD

def country_dashboard(request, pk):
    country = get_object_or_404(Country, pk=pk)
    indicators = Indicator.objects.filter(subcomponent__subcomponent_indicator__country_code=country.cca3).select_related('subcomponent')
    api_data = ApiExtract.objects.filter(country_code=country.cca3)

    context = {
        'country': country,
        'indicators': indicators,
        'api_data': api_data,
    }
    return render(request, 'mytvddata/pages/dashboard/country_profile.html', context)



def export_country_word(request, pk):
    country = get_object_or_404(Country, pk=pk)
    indicators = Indicator.objects.filter(subcomponent__subcomponent_indicator__country_code=country.cca3)
    api_data = ApiExtract.objects.filter(country_code=country.cca3)

    doc = Document()
    doc.add_heading(f"Profil du pays : {country.name}", level=1)
    doc.add_paragraph(f"Capitale : {country.capital}")
    doc.add_paragraph(f"Population : {country.population}")
    doc.add_paragraph(f"Langues : {country.languages}")

    doc.add_heading("Indicateurs", level=2)
    for data in api_data:
        doc.add_paragraph(f"{data.indicator_name} ({data.indicator_code}) - {data.numeric_value} en {data.date}")

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    response["Content-Disposition"] = f'attachment; filename="{country.name}_profile.docx"'
    doc.save(response)
    return response





def preview_json(request):
    """
    Vue pour prévisualiser le fichier JSON uploadé et proposer l'import en base.
    """
    preview_data = None

    if request.method == "POST" and request.FILES.get("json_file"):
        json_file = request.FILES["json_file"]
        try:
            preview_data = json.load(json_file)  # Charger le contenu JSON
            request.session["preview_data"] = preview_data  # stocker en session
            messages.success(request, "Fichier JSON chargé avec succès. Voici la prévisualisation.")
        except Exception as e:
            messages.error(request, f"Erreur lors du chargement du fichier JSON : {e}")

    return render(request, "mytvddata/pages/country/preview_json.html", {"preview_data": preview_data})


def import_json(request):
    """
    Vue pour importer en base le JSON prévisualisé.
    """
    preview_data = request.session.get("preview_data")

    if preview_data:
        try:
            for c in preview_data:
                Country.objects.update_or_create(
                    cca3=c.get("cca3"),
                    defaults={
                        "flag": c.get("flag"),
                        "name": c.get("name"),
                        "official": c.get("official"),
                        "capital": ", ".join(eval(c.get("capital"))) if c.get("capital") else "",
                        "subregion": c.get("subregion"),
                        "area": c.get("area"),
                        "population": c.get("population"),
                        "languages": c.get("languages"),
                        "maps": c.get("data_source"),  # provisoire
                        "ref_data": c.get("ref_data"),
                        "data_source": c.get("data_source"),
                        "country_class": c.get("country_class"),
                    }
                )
            messages.success(request, "Importation réussie ! Les pays ont été ajoutés à la base.")
        except Exception as e:
            messages.error(request, f"Erreur lors de l’importation : {e}")
    else:
        messages.error(request, "Aucune donnée JSON à importer.")

    return redirect("preview_json")

