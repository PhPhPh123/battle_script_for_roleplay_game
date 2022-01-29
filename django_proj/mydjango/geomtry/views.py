from django.shortcuts import render
from django.http import HttpResponse


def get_rectangle_area(request, width, height):
    return HttpResponse(f"Площадь прямоугольника размером {width} на {height} равна {width*height}")


