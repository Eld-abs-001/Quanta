from django.urls import path
from django.contrib.auth import views as auth_views
from . import views


urlpatterns = [
    path('', views.upload_view, name='upload'),
    path('preview/', views.preview_view, name='preview'),
    path('preview/submit/', views.preview_submit_view, name='preview_submit'),
    path('login/', auth_views.LoginView.as_view(template_name='work/login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
]
