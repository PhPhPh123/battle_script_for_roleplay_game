from django.shortcuts import render
from django.http import HttpResponse, HttpResponseRedirect, HttpResponseNotFound

week = {
    'monday': "понедельник",
    'tuesday': "вторник",
    'wednesday': "среда",
    'thursday': "четверт",
    'friday': "пятница",
    "saturday": "суббота",
    "sunday": "воскресенье"

}
def get_rectangle_area(request, width, height):
    return HttpResponse(f"Площадь прямоугольника размером {width} на {height} равна {width * height}")


def week_days(request, days):
    days = days.lower()
    if week.get(days):
        print(week.get(days))
        return HttpResponse(f'День недели -{week[days]}')


def week_days_redirect(request, days):
    list_week_days = list(week)
    if days > len(list_week_days):
        return HttpResponseNotFound(f'Дня недели под номером {days} - не существует')
    thisday = list_week_days[days-1].lower()
    print(thisday)
    return HttpResponseRedirect(f"{thisday}")
