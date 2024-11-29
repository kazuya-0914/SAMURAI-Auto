from django.urls import path
from . import views
from .views import Chap10, Chap11, Chap12

urlpatterns = [
    path('', views.index, name='openpyxl_pandas'),
    path('chap10/', Chap10.as_view(), name='chap10'),
    path('chap10/excel_read/', Chap10.as_view(), name='excel_read_10'),
    path('chap10/excel_write/', Chap10.as_view(), name='excel_write_10'),
    path('chap10/excel_update/', Chap10.as_view(), name='excel_update_10'),
    path('chap10/excel_delete/', Chap10.as_view(), name='excel_delete_10'),   
    path('chap11/', Chap11.as_view(), name='chap11'),
    path('chap11/excel_read/', Chap11.as_view(), name='excel_read_11'),
    path('chap11/excel_select/', Chap11.as_view(), name='excel_select'),
    path('chap11/expression/', Chap11.as_view(), name='expression'),
    path('chap11/connection/', Chap11.as_view(), name='connection'),
    path('chap12/', Chap12.as_view(), name='chap12'),
    path('chap12/work-1-2/', Chap12.as_view(), name='work-1-2'),
    path('chap12/work-3-4/', Chap12.as_view(), name='work-3-4'),
]