"""
URL configuration for TVDdata project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/6.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""

from django.conf import settings
from django.conf.urls.i18n import i18n_patterns
from django.views.static import serve  ## SO HIDE
from django.contrib import admin
from django.urls import path, include,  re_path
from django.conf.urls.static import static
from django.conf.urls import handler404, handler500, handler403

urlpatterns = [
    path('admin/', admin.site.urls),
   # re_path(r'^static/(?P<path>.*)$', serve, {'document_root': settings.STATIC_ROOT}),
  #  path('admin/', admin.site.urls),
    path('i18n/', include('django.conf.urls.i18n')),  # ✅ ajoute le switch de langue
    path('', include('mytvddata.urls'))  
]

handler500 = 'mytvddata.views.custom_500'
handler404="mytvddata.views.handel404"
handler403="mytvddata.views.handel403"


urlpatterns += i18n_patterns(
    path('', include('mytvddata.urls')),  # tes urls principales
)
