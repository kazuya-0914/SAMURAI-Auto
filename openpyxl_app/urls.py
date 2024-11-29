from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='openpyxl'),
    path('save/', views.save, name='save'),
    path('update/', views.update, name='update'),
    path('create/', views.create, name='create'),
    path('change/color/', views.change, name='change-color'),
    path('change/font/', views.change, name='change-font'),
    path('change/border/', views.change, name='change-border'),
    path('join', views.join, name='join'),
    path('cancell', views.cancell, name='cancell'),
    path('insert', views.insert, name='insert'),
    path('delete_rows', views.delete_rows, name='delete_rows'),
    path('append', views.append, name='append'),
    path('delete_cols', views.delete_cols, name='delete_cols'),
    path('backup', views.backup, name='backup'),
    path('prepare', views.prepare, name='prepare'),
    path('line', views.graph, name='line'),
    path('bar', views.graph, name='bar'),
    path('sales', views.sales, name='sales'),
    path('sales_bar', views.sales_bar, name='sales_bar'),
]