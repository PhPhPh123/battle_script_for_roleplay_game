from django.shortcuts import render

from django.http import HttpResponse, HttpResponseNotFound, HttpResponseRedirect

sh = {
    'aries': 'Овен',
    'taurus': 'Телец',
    'twins': 'Близнецы',
    'crayfish': 'Рак',
    'leo': 'Лев',
    'virgo': 'Дева',
    'scales': 'Весы',
    'scorpion': 'Скорпион',
    'ophiuchus': 'Змееносец',
    'sagittarius': 'Стрелец',
    'capricorn': 'Козерог',
    'aquarius': 'Водолей',
}


def get_info_sign_horoscope(requests, sign_horoscope):
    if sh.get(sign_horoscope):
        return HttpResponse(f'Знак зодиака - {sh[sign_horoscope]}')
    else:
        return HttpResponseNotFound(f'Знак зодиака "{sign_horoscope}" не определен.')


def get_info_sign_horoscope_by_number(requests, sign_horoscope: int):
    zodiacs = list(sh)
    if sign_horoscope > len(zodiacs):
        return HttpResponseNotFound(f"Знака {sign_horoscope}- не существует")
    name_zodiac = zodiacs[sign_horoscope-1]
    return HttpResponseRedirect(f"/{name_zodiac}/")
