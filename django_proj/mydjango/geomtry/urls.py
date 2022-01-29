from django.urls import path
from . import views

urlpatterns = [
    path("rectangle/<int:width>/<int:height>", views.get_rectangle_area)

]

