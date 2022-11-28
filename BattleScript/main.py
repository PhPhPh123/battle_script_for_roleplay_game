"""
Основной модуль скрипта, в котором производятся все расчеты.
Его работа основана на парсинге информации с excel-файла и создание из него списка с юнитами первого и второго войска
Затем эти списки отправляются в первую стадию боя(first_stage_battle)где наносят друг другу урон согласно их параметра
атаки, затем во вторую фазу боя(second_stage_battle), где идет сравнения их боевой эффективности(он же уровень здоровья)
"""

import openpyxl as op
from random import shuffle, randint
from copy import deepcopy


class Parser:
    """
    Данный класс занимается парсингом данных из excel-файла и на выходе дает 2 полностью готовые для проведения битвы
    армии
    """
    def __init__(self, xls_worksheet):
        # основные ключи, необходимые для формирования юнитов в обоих армиях
        self.keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                                   "damage", "defence", "morale", "targets", "combat_bonus", "amount")
        self.xls_worksheet = xls_worksheet

    def parse_excel(self, worksheet, minrow: int, maxrow: int, mincol: int, maxcol: int) -> list[dict, ...]:
        """
        Основной метод, парсящий данные из excel-страницы
        :param worksheet: рабочий лист excel
        :param minrow: строка, с которой начинается парсинг
        :param maxrow: строка, которой заканчивается парсинг
        :param mincol: колонка, с которой начинается парсинг
        :param maxcol: колонка, которой заканчивается парсинг
        :return: отпаршенный лист со словарями
        """
        pars_list = []  # инициализация пустого листа который впоследствии станет списком словарей

        for row in worksheet.iter_rows(min_row=minrow, max_row=maxrow, min_col=mincol, max_col=maxcol,
                                       values_only=True):
            if not row[0]:  # если парсер доходит до пустой строки, значит юниты кончились и парсинг останавливается
                break
            army_or_ability_dict = dict(zip(self.keys_for_army_dict, row))  # отпаршенные значения связываются с ключами
            pars_list.append(army_or_ability_dict)  # словари добавляются в общий список

        return pars_list

    def create_army_dicts(self, parslist: list[dict, ...]) -> list[dict, ...]:
        """
        Данная функция разбивает список словарей, в котором одинаковые юниты собраны в один стак на отдельные юниты
        которые затем будут пронумерованы в методе numerate_army_units
        :param parslist: лист со словарями в котором одиноковые юниты собраны в кучу
        :return: лист в котором каждому юниту соответствует один отдельный словарь
        """
        army_list = []  # инициализация списка словарей

        # иду по списку стаков юнитов
        for i in parslist:

            # создание отдельных юнитов продолжается только когда их количество в стаке больше или равно 1, если
            # при достижении 0 разбитие стака на отдельные юниты прекращается
            while i["amount"] >= 1:

                # в список отдельных юнитов добавляются словари со всеми ключ-значениями кроме последнего,
                # т.к. он отвечает за количество юнитов в стаке [:-1]
                army_list.append(dict(list(i.items())[:-1]))
                i["amount"] -= 1  # уменьшаю значение стака на 1

        numerated_army_list = self.numerate_army_units(army_list)  # отправляю список с отдельными юнитами на нумерацию

        return numerated_army_list

    @staticmethod
    def numerate_army_units(arm_list: list[dict, ...]) -> list[dict, ...]:
        """
        Данный метод добавляет каждому отдельному юниту его номерной знак, чтобы они были отличны друг от друга
        :param arm_list: список со словарями, в котором юниты сепарированы друг от друга
        :return: список со словарями, в котором юниты сепарированы друг от друга и имеют уникальные номера
        """
        # список словарей сортируется по именам юнитов
        sorted_arm_list = sorted(arm_list, key=lambda islist: islist["unit_name"])

        # нумерация отсортированного списка и придание каждому юниту его номера
        for num, unit in enumerate(sorted_arm_list, start=1):
            unit["unit_name"] += f" ,юнит №{num}"

        numerated_arm_list = deepcopy(sorted_arm_list)  # копирую список глубоким копированием

        return numerated_arm_list

    def create_both_armies(self):
        """
        Метод, управляющий созданием армий
        :return: списки со словарями обоих армий
        """
        # создание отпаршенных списков со словарями в которых юниты стекированы, ключевые аргументы отвечают за
        # строки и столбцы, в пределах которых будет идти парсинг из excel-листа
        first_army_pars_list = self.parse_excel(self.xls_worksheet, minrow=4, maxrow=35, mincol=2, maxcol=11)
        second_army_pars_list = self.parse_excel(self.xls_worksheet, minrow=4, maxrow=35, mincol=14, maxcol=23)

        # создание списков со словарями, в которых юниты отделены от стеков
        first_army_combat_list = self.create_army_dicts(first_army_pars_list)
        second_army_combat_list = self.create_army_dicts(second_army_pars_list)

        # армии прогоняются через функцию, которая добавляет юнитам ключ-значение нехватки их боевой эффективности
        # относительно максимальной(дефакто юнит ранен ДО боя и начинает бой раненым)
        first_army_with_shortage = count_shartage_before_and_after_battle(first_army_combat_list)
        second_army_with_shortage = count_shartage_before_and_after_battle(second_army_combat_list)

        return first_army_with_shortage, second_army_with_shortage


class FirstBattleStage:
    @staticmethod
    def count_battle(attacker_list: list[dict, ...], defender_list: list[dict, ...]) -> tuple:
        """
        Данный метод проводит первый этап боя. Его смысл в том, что каждый юнит атакующей армии должен попытаться
        нанести урон юниту защищающейся армии, если тип юнита защищающейся армии уязвим для типа юнита атакующей армии
        :return: списки со словарями защищающейся армии и ее кладбища
        """

        for attacker_army_unit in attacker_list:  # прохожу по каждому юниту атакующей армии

            for defender_army_unit in defender_list:  # затем прохожу по каждому юниту защищающейся армии

                # если у атакующего юнита есть возможность наносить урон(targets содержит значения)
                if attacker_army_unit["targets"]:

                    # если защищающийся юнит узявим к типу атакующего юнита и у защищающегося юнита больше 0 текущей
                    # комбатки(здоровья) то ему наносится урон иначе ищется другой юнит в защищающейся армии
                    if defender_army_unit["unit_type"] in attacker_army_unit["targets"] and defender_army_unit[
                        "current_combat"] > 0:

                        # данная формула считает урон
                        damage = defender_army_unit["max_combat"] // 100 * (
                                attacker_army_unit["damage"] - defender_army_unit["defence"])

                        # если защита обороняющегося юнита больше чем атака нападающего, то damage может стать
                        # отрицательным, что не нужно, поэтому, изменения "current_combat" не произойдут вообще
                        if damage > 0:

                            # уменьшаю текущую комбатку юнита на величину урона
                            defender_army_unit["current_combat"] -= damage

                        shuffle(defender_list)  # перетасовываю список юнитов, чтобы был эффект случайности и одни
                        # юниты могли получать урон более чем один раз

                        break  # этот слом цикла нужен, чтобы предотвратить насение более чем одного удара одним
                        # юнитом атакующей армии, если он успешно атаковал кого то(даже если защитник проигнорировал
                        # его атаку своей защитой) то больше атаковать не может. break прекращает итерацию по списку
                        # защищающийся юнитов и возвращает исполнение в итерацию по атакующим юнитам

        # отправляю списки в функцию, которая очищает их от мертвых юнитов и создает отдельных список со словарями
        # с мертвыми юнитами(юнитами у которых "current_combat" стала меньше или равно 0)
        defender_after_casualties, defender_graveyard = remove_dead_units(deepcopy(defender_list))

        return defender_after_casualties, defender_graveyard


class SecondBattleStage:
    MORALE_MODIFIER = 2  # константа множителя морали
    BASE_CASUALTIES = 30  # константа базового уровня потерь

    def __init__(self, first_army: list[dict, ...], second_army: list[dict, ...],
                 first_army_first_stage_graveyard: list[dict, ...], second_army_first_stage_graveyard: list[dict, ...]):
        self.second_stage_advantage = 0  # инициализация преемущества одной армии на другой, в процентах
        self.winner = None  # инициализация победителя по итогам второй стадии
        self.first_army = first_army
        self.second_army = second_army

        # кладбища от первой стадии битвы
        self.first_army_first_stage_graveyard = first_army_first_stage_graveyard
        self.second_army_first_stage_graveyard = second_army_first_stage_graveyard

        # кладбища от второй стадии битвы, инициализация
        self.first_army_second_stage_graveyard = None
        self.second_army_second_stage_graveyard = None

    def count_combat_points(self, army: list[dict, ...]) -> int:
        """
        Данный метод подсчитывает общее количество очков комбатки(здоровья) армии, в подсчете также участвует
        модификатор морали, который может удвоить уменьшить в 2 раза размер комбатки юнита. Мораль это буквально
        процентный шанс того, что произойдет положительное или отрицательное изменение согласно модулю цифры комбатки
        Например есть мораль -15% то это значит, что с 15% шансом юнит будет задизморален
        :return: общая сумма комбатки армии
        """
        army_sum_combat = 0  # инициализация суммы

        for unit in army:  # прохожу юнитом по армии
            if unit["current_combat"] > 0:  # если юнит живой

                if 0 < unit["morale"] >= randint(0, 100):  # если его мораль позитивная и значение случайности удачно
                    # добавляю в общий список комбатку юнита, умноженные на мод. морали
                    army_sum_combat += unit["current_combat"] * self.MORALE_MODIFIER
                    unit["morale_boost"] = 'good'  # вешаю юниту флаг, что он воодушевился в бою и увеличил свою мораль

                elif 0 > unit["morale"] <= randint(-100, 0):  # если его мораль негативная и значение случайности удачно
                    # добавляю в общий список комбатку юнита, поделенный на мод. морали
                    army_sum_combat += int(unit["current_combat"] / self.MORALE_MODIFIER)
                    unit["morale_boost"] = 'bad'  # вешаю юниту флаг, что он пал духом в бою и уменьшил свою мораль
                else:  # если не произошло событий морали, то просто добавляется в сумму обычная комбатка юнита
                    army_sum_combat += unit["current_combat"]
                    unit["morale_boost"] = False  # Вешается флаг, что событий морали с юнитом не было
        return army_sum_combat

    def count_army_advantage(self, first_army_sum_combat: int, second_army_sum_combat: int) -> None:
        """
        Данный метод посчитывает процент преимущества первой армии на второй. Если преимущество выше 0, то оно за
        первой армией, если ниже - за второй
        :param first_army_sum_combat: сумма комбатки первой армии
        :param second_army_sum_combat: сумма комбатки второй армии
        :return: None
        """
        advantage = int((first_army_sum_combat - second_army_sum_combat) / min(first_army_sum_combat, second_army_sum_combat) * 100 / 5)
        self.second_stage_advantage = advantage

    def count_second_stage(self, first_army: list[dict, ...], second_army: list[dict, ...]) -> tuple:
        """
        Данный метод считает итоги проведения второй стадии боя
        :param first_army: список словарей первой армии
        :param second_army: список словарей второй армии
        :return: списки словарей первой и второй армии по итогам боя
        """
        # подсчет суммы очков комбатки(cp) обеих армий
        first_army_final_cp = self.count_combat_points(first_army)
        second_army_final_cp = self.count_combat_points(second_army)

        # если подсчет комбатки одной из армий выдал сумму в 0, значит одна из этих армий полностью уничтожена
        # в первой стадии и метод возвращает результаты обоих армий после очищения ее от трупов соответствующей функцией
        # remove_dead_units
        if first_army_final_cp == 0 or second_army_final_cp == 0:
            return remove_dead_units(first_army), remove_dead_units(second_army)

        # подсчет преимущества первой армии над второй
        self.count_army_advantage(first_army_final_cp, second_army_final_cp)

        # на основе полученного преимущества высчитываются потери первой и второй армии
        if self.second_stage_advantage > 0:  # если преимущество на стороне первой армии

            # то ее базовые потери уменьшаются на модификатор преимущества деленного на 2
            first_army_casualties = self.BASE_CASUALTIES - self.second_stage_advantage // 2
            # а преимущество первой армии добавляется к потерям второй армии
            second_army_casualties = self.BASE_CASUALTIES + self.second_stage_advantage

        elif self.second_stage_advantage < 0:  # если преимущество на стороне второй армии
            # то из базовых потерь вычитается модуль преемущества(т.е. значение отрицательное) деленное на 2
            second_army_casualties = self.BASE_CASUALTIES - abs(self.second_stage_advantage // 2)
            # а преемущество второй армии добавляется к потерям первой армии
            first_army_casualties = self.BASE_CASUALTIES + abs(self.second_stage_advantage)
        else:  # если преемущество равно 0, то это фактически ничья и обе армии получаются чистые базовые потери
            first_army_casualties = self.BASE_CASUALTIES
            second_army_casualties = self.BASE_CASUALTIES

        # если потери одной из армий выходят за пределы (0, 100), то значения корректируются согласно диапазону
        # это нужно, чтобы избежать багов в отрицательными потерями или потерями, превышающими 100%
        if first_army_casualties >= 100:
            first_army_casualties = 100
        if first_army_casualties <= 0:
            first_army_casualties = 0
        if second_army_casualties >= 100:
            second_army_casualties = 100
        if second_army_casualties <= 0:
            second_army_casualties = 0

        # установка флага победителя
        if first_army_final_cp > second_army_final_cp:
            self.winner = "победа первой армии"
        elif first_army_final_cp < second_army_final_cp:
            self.winner = "победа второй армии"
        else:
            self.winner = "Ничья"

        # применение процента потерь к каждому отдельному юниту в обеих армиях
        for unit in first_army:
            unit["current_combat"] = int(unit["current_combat"] / 100 * (100 - first_army_casualties))
        for unit in second_army:
            unit["current_combat"] = int(unit["current_combat"] / 100 * (100 - second_army_casualties))

        # очищение армий от трупов и создание кладбища второй стадии боя
        first_army_after_fight, self.first_army_second_stage_graveyard = remove_dead_units(first_army)
        second_army_after_fight, self.second_army_second_stage_graveyard = remove_dead_units(second_army)

        return first_army_after_fight, second_army_after_fight

    def control_battle_logic(self):
        """
        Управляющий метод второй стадии боя, вызывающий необходимые боевые и послебоевые методы
        :return: списки обоих армий, а также и итоговые кладбища
        """
        # Вызов метода, подсчитывающего основные результаты второй стадии и возвращающий списки армий
        first_army_after_fight, second_army_after_fight = self.count_second_stage(self.first_army, self.second_army)

        # Вызов методов, подсчитывающего итоговые кладбища по результатам первой и второй стадии битвы, списки словарей
        # обоих кладбища складываются
        first_army_final_graveyard = remove_negative_casualties_from_graveyard(self.first_army_first_stage_graveyard +
                                                                               self.first_army_second_stage_graveyard)
        second_army_final_graveyard = remove_negative_casualties_from_graveyard(self.second_army_first_stage_graveyard +
                                                                                self.second_army_second_stage_graveyard)

        # итоговые списки словарей армий пройденных через функцию подсчета их потерь по итогам боя
        self.first_army = count_shartage_before_and_after_battle(first_army_after_fight)
        self.second_army = count_shartage_before_and_after_battle(second_army_after_fight)

        return self.first_army, self.second_army, first_army_final_graveyard, second_army_final_graveyard


def remove_dead_units(unitlist: list[dict, ...]) -> tuple[list[dict, ...], list[dict, ...]]:
    """
    Данная отдельная функция удаляет из списка основной армии погибшие юниты и создает отдельный список кладбища с ними
    :param unitlist: армия, которую нужно проверить на трупы
    :return: два списка со словарями, один действующая армия, другой - кладбище
    """
    new_unit_list = []  # инициализация нового списка живой армии
    graveyard_list = []  # инициализация кладбища

    for unit in unitlist:
        if unit["current_combat"] > 0:  # если комбатка юнита больше нуля, то он живой и отправляется в живую армию
            new_unit_list.append(unit)
        else:  # если комбатка 0 или меньше, то он отправляется на кладбище
            graveyard_list.append(unit)

    return new_unit_list, graveyard_list


def count_shartage_before_and_after_battle(list_army: list[dict, ...]) -> list[dict, ...]:
    """
    Данная функция определяет в войске нехватку боевой эффективности(комбатки) относительно максимальной эффективности.
    Начальная нехватка ДО боя определяется ключем "shortage", а после боя ключом "casualties". Т.е. "shortage" это то,
    те потери, которые были у юнита до сражения, а "casualties" это потери, которые возникли в результате сражения, но
    без учета изначальной нехватки
    :param list_army: список армии со словарями юнитов
    :return: лист с добавленными ключ-значениями нехватки и потерь
    """
    for unit in list_army:
        if unit['unit_type'] not in 'ч+п+о':  # Данные теги юнитов являются одиночными и не имеют статуса нехватки
            if "shortage" not in unit:  # перед битвой ключа потерь еще нет и его нужно создать и посчитать значения
                # нехватки комбатки
                unit["shortage"] = int(100 - (unit["current_combat"] / unit['max_combat'] * 100))
            else:  # если ключ нехватки уже есть, значит нужно подсчитать потери на основе урона в бою и нехватки до боя
                unit["casualties"] = int(100 - (unit["current_combat"] / unit['max_combat'] * 100)) - unit["shortage"]
        else:  # для одиночных юнитов по умолчанию ставится потери равные 0 т.к. их убийство произойдет, по механике,
            # только в случае полного уничтожения армии(вайпа), в иных случаях они выживают
            unit["casualties"] = 0
    return list_army


def remove_negative_casualties_from_graveyard(graveyard_list: list[dict, ...]) -> list[dict, ...]:
    """
    Данная функция приводит в порядок данные по кладбищу, чтобы были видны реальные потери casualties без учета
    возможных аномалий связанных с отрицательными значениями
    :param graveyard_list: список словарей с мёртвыми юнитами
    :return: список словарей с мертвыми юнитами и правильными данными их потерь
    """

    for corpse in graveyard_list:
        if corpse['unit_type'] in 'ч+п+о':  # попадание в кладбище означает, что одиночный юнит умер и потери равны 1
            corpse["casualties"] = 1
        else:  # если юнит не одиночный, то его реальные потери это 100% - процентная нехватка до боя(shortage)
            corpse["casualties"] = int(100 - corpse['shortage'])
    return graveyard_list


def write_txt_file(first_army: list[dict, ...], second_army: list[dict, ...],
                   first_army_graveyard: list[dict, ...], second_army_graveyard: list[dict, ...],
                   second_stage_obj: SecondBattleStage) -> None:
    """
    Данная функция записывает в текстовый файл результаты боя и вызывает метод статистики, который дозаписывает
    в файл статистическую информацию по результатам боя
    :param first_army: список словарей первой армии после боя
    :param second_army: список словарей второй армии после боя
    :param first_army_graveyard: список словарей кладбища первой армии
    :param second_army_graveyard: список словарей кладбища второй армии
    :param second_stage_obj: объект второй стадии боя
    :return: None, результат это запись в файл
    """
    with open("result.txt", mode='w', encoding="utf-8") as res:
        # записывает кто победит по результатам сражения на основе флага победителя
        res.write(f'Бой проведен. Его результатом стала {second_stage_obj.winner}')
        # записывает процент преемущества победителя во второй стадии боя
        res.write(f'\nПреемущество победителя во второй фазе боя - {abs(second_stage_obj.second_stage_advantage)}%\n')

    # вызов методов статистики для первой и второй армии, магические числа 1 и 2 идут как флаги первой и второй армии
    write_army_statistics(first_army, first_army_graveyard, 1)
    write_army_statistics(second_army, second_army_graveyard, 2)


def write_army_statistics(army_list: list[dict, ...], grave_list: list[dict, ...], number_of_army: int) -> None:
    """
    Данный метод записывает в текстовый файл статистику по результатам боя для обеих армий
    :param army_list: список словарей армии
    :param grave_list: список словарей кладбища
    :param number_of_army: номер армии
    :return: None, результат это запись в файл
    """

    with open("result.txt", mode='a', encoding="utf-8") as stat:
        # строка с названием того, чья статистика будет записана
        if number_of_army == 1:
            stat.write("\nСтатистика первой армии:")
        elif number_of_army == 2:
            stat.write("\nСтатистика второй армии:")

        if army_list:  # если в армии есть хотябы один юнит

            stat.write("\nВыжившие юниты:\n")
            for line in army_list:  # иду юнитом по списку армии

                # записываю имя юнита, его оставшуюся комбатку, а также то, сколько людей(эльфов, гномов, другое)
                stat.write(f'{line["unit_name"]} оставшаяся комбатка: {line["current_combat"]} '
                           f'погибло людей: {line["casualties"]} \n')
        else:  # если юнитов нет, то армия уничтожена
            stat.write("\nАрмия полностью уничтожена")

        # иду по кладблищу и записываю имена погибших юнитов
        stat.write('\nПогибшие юниты:\n')
        for corpse in grave_list:
            stat.write(f'{corpse["unit_name"]}\n')
        # если кладбище пустое, то пишу об этом
        if not grave_list:
            stat.write('Ни один юнит не уничтожен полностью')

        # записываю итоговые суммарные потери по населению, по итогам битвы собранные из юнитов из кладбища и армии
        stat.write(f'\nCуммарные потери по количеству населения равны: '
                   f'{sum(i["casualties"] for i in grave_list + army_list)}\n\n')  # генератор подсчета суммы

        # блок подсчета морали юнитов
        stat.write('Бафы и дебафы морали:\n')

        morale_units = []  # инициализация списка с юнитами, прошедшими проверку морали

        for unit in army_list + grave_list:  # иду по списку словарей армии и кладбища
            if unit["morale_boost"] == 'good':  # если мораль юнита бафнулась, то делаю об этом запись
                stat.write(f'{unit["unit_name"]} воодушевился и удвоил свою комбатку в бою\n')
                morale_units.append(unit)  # добавляю юнита в список юнитов, прошедших проверку морали

            elif unit["morale_boost"] == 'bad':  # если мораль юнита дебафнулась, то делаю об этом запись
                stat.write(f'{unit["unit_name"]} пошатнулся боевым духом и струсил\n')
                morale_units.append(unit)  # добавляю юнита в список юнитов, прошедших проверку морали
            elif not unit["morale_boost"]:  # Если мораль юнита = False, значит он проверок морали не прошел и запись
                # игнорируется
                pass
        # Если список юнитов с моралью пустой, значит никто себя не проявил и нужно об этом сделать запись
        if not morale_units:
            stat.write('Никто не проявил себя особым образом')
        stat.write('\n\n')  # перебрасываю каретку на 2 строки, чтобы отделить статистику разных армий друг от друга


def main_logic() -> None:
    """
    Это основная фасадная функция, контролирующая работу скрипта, по этапам
    :return: None
    """

    """Создание рабочей книги и рабочего листа excel"""
    xls_workbook = op.load_workbook("table.xlsx", read_only=True)
    xls_worksheet = xls_workbook.active

    """Этап парсинга данных из excel-файла и создание на его основе двух армий в виде двух списков со словарями"""
    parser_obj = Parser(xls_worksheet)
    first_army_combat_list, second_army_combat_list = parser_obj.create_both_armies()

    """Первая стадия битвы"""
    first_battle_obj = FirstBattleStage()
    # Сначала первая армия наносит урон, а вторая получает
    first_army_first_stage, first_army_graveyard_1st_stage = first_battle_obj.count_battle(second_army_combat_list,
                                                                                           first_army_combat_list)
    # Затем вторая армия наносит урон, а первая получает
    second_army_first_stage, second_army_graveyard_1st_stage = first_battle_obj.count_battle(first_army_combat_list,
                                                                                             second_army_combat_list)
    """Вторая стадия битвы"""
    second_stage_obj = SecondBattleStage(first_army_first_stage, second_army_first_stage,
                                         first_army_graveyard_1st_stage, second_army_graveyard_1st_stage)

    (first_army_second_stage, second_army_second_stage,
     first_army_graveyard, second_army_graveyard) = second_stage_obj.control_battle_logic()

    """Стадия после боя"""
    write_txt_file(first_army_second_stage, second_army_second_stage,
                   first_army_graveyard, second_army_graveyard, second_stage_obj)


if __name__ == "__main__":
    main_logic()
