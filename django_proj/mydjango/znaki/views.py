from django.shortcuts import render

from django.http import HttpResponse


def vodoley(request):
    return HttpResponse("Это водолей, это я")


def ribi(request):
    return HttpResponse("А это моя сестра")

