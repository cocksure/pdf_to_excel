from django.urls import path

from converter import views

urlpatterns = [
    path('', views.upload_pdf, name='upload_pdf'),

]