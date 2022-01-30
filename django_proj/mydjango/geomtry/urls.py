from django.urls import path
from . import views

urlpatterns = [
    path("rectangle/<int:width>/<int:height>", views.get_rectangle_area),
    path("rectangle/<int:width>/<int:height>", views.get_rectangle_area),
    path("weekdays/<int:days>", views.week_days_redirect),
    path("weekdays/<days>", views.week_days)
]

