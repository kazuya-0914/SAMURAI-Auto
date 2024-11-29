from django.urls import path
from . import views
from .views import Chap07, Chap08, Chap09

urlpatterns = [
    path('', views.index, name='pandas'),
    path('chap07/', Chap07.as_view(), name='chap07'),
    path('chap07/csv_read/', Chap07.as_view(), name='csv_read'),
    path('chap07/csv_write/', Chap07.as_view(), name='csv_write'),
    path('chap07/excel_read/', Chap07.as_view(), name='excel_read'),
    path('chap07/excel_write/', Chap07.as_view(), name='excel_write'),
    path('chap07/data_filter/', Chap07.as_view(), name='data_filter'),
    path('chap07/data_mean/', Chap07.as_view(), name='data_mean'),
    path('chap07/data_result/', Chap07.as_view(), name='data_result'),
    path('chap08/', Chap08.as_view(), name='chap08'),
    path('chap08/data_iloc/', Chap08.as_view(), name='data_iloc'),
    path('chap08/to_numeric/', Chap08.as_view(), name='to_numeric'),
    path('chap08/to_string/', Chap08.as_view(), name='to_string'),
    path('chap08/loc/', Chap08.as_view(), name='loc'),
    path('chap08/concat/<int:axis>', Chap08.as_view(), name='concat'),
    path('chap09/', Chap09.as_view(), name='chap09'),
    path('chap09/practice/', Chap09.as_view(), name='practice'),
]