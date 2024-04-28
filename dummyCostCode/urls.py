from django.urls import path
from . import views
app_name = "dummyCostCode"
urlpatterns = [
    path('', views.post_list, name='post_list'),
    path('sampleResultsAllDates/', views.sampleResultsAllDates, name='sampleResultsAllDates'),
    path('calculateTotalBillingAmountDetail/', views.calculateTotalBillingAmountDetail, name='calculateTotalBillingAmountDetail'),
    path('compSumAndDetail/', views.compSumAndDetail, name='compSumAndDetail'),
    path('compPrevDetChangesAndCurDetAll.html/', views.compPrevDetChangesAndCurDetAll, name='compPrevDetChangesAndCurDetAll'),
    path('finalReport.html/', views.finalReport, name='finalReport'),
    
    
]