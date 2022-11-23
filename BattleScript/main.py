"""
Основной модуль скрипта, в котором производятся все расчеты.
Его работа основана на парсинге информации с excel-файла и создание из него списка с юнитами первого и второго войска
Затем эти списки отправляются в первую стадию боя(first_stage_battle)где наносят друг другу урон согласно их параметра
атаки, затем во вторую фазу боя(second_stage_battle), где идет сравнения их боевой эффективности(он же уровень здоровья)

"""

import openpyxl as op
from typing import Any
from random import shuffle, randint
from copy import deepcopy

"Глобальные переменные и константы"
MORALE_MODIFIER = 2  # Глобальная константа морали
BASE_CASUALTIES = 30  # Глобальная константа базового уровня потерь
winner = None  # инициализация глобальной переменная победителя боя
second_stage_advantage = 0  # инициализация глобальная переменной преимущества победителя во второй фазе боя


def first_stage_battle(attacker_list, defender_list):
    """
    this function counts first stage battle and modified army lists to use it in second battle stage
    :return: None
    """

    for attacker_army_unit in attacker_list:  # iterate by unit in first army list
        for defender_army_unit in defender_list:  # iterate by unit in second army list
            # if unit in second army have type which damaged by first army unit and second army have
            # at least 1 combat point(HP)
            if attacker_army_unit["targets"]:
                if defender_army_unit["unit_type"] in attacker_army_unit["targets"] and defender_army_unit[
                    "current_combat"] > 0:

                    # this formula counts damage
                    dmg1 = defender_army_unit["max_combat"] // 100 * (
                            attacker_army_unit["damage"] - defender_army_unit["defence"])

                    if dmg1 > 0:  # if protection bigger then the damage then no damage is dealt(it needs to prevent
                        # negative damage)
                        defender_army_unit["current_combat"] -= dmg1  # dealt damage to unit

                    shuffle(defender_list)  # random function which shuffle army list in each iteration
                    #  This is necessary to introduce an element of chance

                    break  # this break in fact return us to first "for" cycle. Its need to prevent hitting more then
                    # 1 unit

    defender_after_casualties, defender_graveyard = remove_dead_units(deepcopy(defender_list))

    return defender_after_casualties, defender_graveyard


def combat_points_counter(army):
    """
    this function counts sum of combat points to use it in second battle stage
    also this function use moral modifier to double (or halve) combat points if unit is lucky
    :return: sum of first and second army combat points
    """
    army_sum_combat = 0  # final combat points of army

    for unit in army:  # iterate by unit in army list
        if unit["current_combat"] > 0:  # if unit is not dead(combat point <= 0)
            if 0 < unit["morale"] >= randint(0, 100):  # positive morale check(morale>0).
                # check succeeds if morale >= randint
                army_sum_combat += unit["current_combat"] * MORALE_MODIFIER  # add combat point * morale to sum
                unit["morale_boost"] = 'good'
            elif 0 > unit["morale"] <= randint(-100, 0):  # negative morale check(morale<0).
                # check succeeds if morale >= randint
                unit["morale_boost"] = 'bad'
                army_sum_combat += int(unit["current_combat"] / MORALE_MODIFIER)  # add combat point // morale to sum
            else:
                army_sum_combat += unit["current_combat"]  # add combat points to sum if morale == 0
                unit["morale_boost"] = False
    return army_sum_combat


def army_advantage(first_cb: int, second_cb: int) -> int:
    """
    this function counts advantage of army with bigger combat points expressed as a percentage
    :param first_cb: first army combat point
    :param second_cb: second army combat points
    :return: advantage of the first army over the second, if advantage < 0 it means first army is weaker them second
    """
    advantage = int((first_cb - second_cb) / min(first_cb, second_cb) * 100 / 5)
    global second_stage_advantage
    second_stage_advantage = advantage
    return advantage


def second_stage_battle(first_army, second_army):
    first_army_final_cp = combat_points_counter(first_army)
    second_army_final_cp = combat_points_counter(second_army)

    if first_army_final_cp == 0 or second_army_final_cp == 0:
        return remove_dead_units(first_army), remove_dead_units(second_army)

    advantage = army_advantage(first_army_final_cp, second_army_final_cp)

    if advantage > 0:
        first_army_casualties = BASE_CASUALTIES - advantage // 2
        second_army_casualties = BASE_CASUALTIES + advantage
    elif advantage < 0:
        first_army_casualties = BASE_CASUALTIES + abs(advantage)
        second_army_casualties = BASE_CASUALTIES - abs(advantage // 2)
    else:
        first_army_casualties = BASE_CASUALTIES
        second_army_casualties = BASE_CASUALTIES

    if first_army_casualties >= 100:
        first_army_casualties = 100
    if first_army_casualties <= 0:
        first_army_casualties = 0
    if second_army_casualties >= 100:
        second_army_casualties = 100
    if second_army_casualties <= 0:
        second_army_casualties = 0

    global winner
    if first_army_final_cp > second_army_final_cp:
        winner = "победа первой армии"
    elif first_army_final_cp < second_army_final_cp:
        winner = "победа второй армии"
    else:
        winner = "Ничья"

    for unit in first_army:
        unit["current_combat"] = int(unit["current_combat"] / 100 * (100 - first_army_casualties))
    for unit in second_army:
        unit["current_combat"] = int(unit["current_combat"] / 100 * (100 - second_army_casualties))

    first_army_afterfight, first_army_graveyard = remove_dead_units(first_army)
    second_army_afterfight, second_army_graveyard = remove_dead_units(second_army)

    return first_army_afterfight, second_army_afterfight, first_army_graveyard, second_army_graveyard


def remove_dead_units(unitlist: list) -> tuple[list[Any], list[Any]]:
    new_list = []
    graveyard_list = []
    for unit in unitlist:
        if unit["current_combat"] > 0:
            new_list.append(unit)
        else:
            graveyard_list.append(unit)

    return new_list, graveyard_list


class Parser:
    def __init__(self, xls_worksheet):
        self.keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                                   "damage", "defence", "morale", "targets", "combat_bonus", "amount")
        self.xls_worksheet = xls_worksheet

    def parse_excel(self, worksheet, minrow, maxrow, mincol, maxcol):
        pars_list = []

        for row in worksheet.iter_rows(min_row=minrow, max_row=maxrow, min_col=mincol, max_col=maxcol,
                                       values_only=True):
            if not row[0]:
                break
            army_or_ability_dict = dict(zip(self.keys_for_army_dict, row))
            pars_list.append(army_or_ability_dict)

        return pars_list

    @staticmethod
    def numerate_army_units(arm_list):
        sorted_arm_list = sorted(arm_list, key=lambda islist: islist["unit_name"])

        for num, unit in enumerate(sorted_arm_list, start=1):
            unit["unit_name"] += f" ,юнит №{num}"
        numerated_arm_list = deepcopy(sorted_arm_list)

        return numerated_arm_list

    def create_army_lists(self, parslist):
        army_list = []

        for i in parslist:
            while i["amount"] >= 1:
                army_list.append(dict(list(i.items())[:-1]))
                i["amount"] -= 1
        numerated_army_list = self.numerate_army_units(army_list)

        return numerated_army_list

    def create_both_armies(self):
        first_army_pars_list = self.parse_excel(self.xls_worksheet, minrow=4, maxrow=35, mincol=2, maxcol=11)
        second_army_pars_list = self.parse_excel(self.xls_worksheet, minrow=4, maxrow=35, mincol=14, maxcol=23)

        first_army_combat_list = self.create_army_lists(first_army_pars_list)
        second_army_combat_list = self.create_army_lists(second_army_pars_list)

        return first_army_combat_list, second_army_combat_list


def shortage_before_and_after_battle(list_army: list[dict, ]) -> list[dict]:
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


def casualties_for_graveyard(graveyard_list: list):
    for corpse in graveyard_list:
        if corpse['unit_type'] in 'ч+п+о':
            corpse["casualties"] = 1
        else:
            corpse["casualties"] = int(100 - corpse['shortage'])
    return graveyard_list


def army_statistics(army_list, grave_list, number_of_army):
    with open("result.txt", mode='a', encoding="utf-8") as stat:
        if number_of_army == 1:
            stat.write("\nСтатистика первой армии:")
        else:
            stat.write("\nСтатистика второй армии:")

        if army_list:
            first_arm_cas = 0
            stat.write("\nВыжившие юниты:\n")
            for line in army_list:
                first_arm_cas += line["casualties"]
                stat.write(f'{line["unit_name"]} оставшаяся комбатка: {line["current_combat"]} '
                           f'погибло юнитов: {line["casualties"]} \n')
        else:
            stat.write("\nАрмия полностью уничтожена")

        stat.write('\nПогибшие юниты:\n')
        for corpse in grave_list:
            stat.write(f'{corpse["unit_name"]}\n')
        if not grave_list:
            stat.write('Ни один юнит не уничтожен полностью')

        stat.write(f'\nCуммарные потери по количеству населения равны: '
                   f'{sum(i["casualties"] for i in grave_list + army_list)}\n\n')

        stat.write('Бафы и дебафы морали:\n')
        for hero in army_list + grave_list:
            try:
                if hero["morale_boost"] == 'good':
                    stat.write(f'{hero["unit_name"]} воодушевился и удвоил свою комбатку в бою\n')
                elif hero["morale_boost"] == 'bad':
                    stat.write(f'{hero["unit_name"]} пошатнулся боевым духом и струсил\n')
            except KeyError:
                pass
        else:
            stat.write('Никто не проявил себя особым образом')
        stat.write('\n\n')


def file_writer(first_list, second_list, first_grave_list, second_grave_list):
    with open("result.txt", mode='w', encoding="utf-8") as res:
        res.write(f'Бой проведен. Его результатом стала {winner}')
        res.write(f'\nПреемущество победителя во второй фазе боя - {abs(second_stage_advantage)}%\n')

    army_statistics(first_list, first_grave_list, 1)
    army_statistics(second_list, second_grave_list, 2)


def main_logic() -> None:
    """
    Это основная фасадная функция, контролирующая работу скрипта, по этапам
    :return: None
    """
    # Создание рабочей криги и рабочего листа excel
    xls_workbook = op.load_workbook("table.xlsx", read_only=True)
    xls_worksheet = xls_workbook.active

    # Этап парсинга данных из excel-файла и создание на его основе двух армий в виде двух списков со словарями
    parser_obj = Parser(xls_worksheet)
    first_army_combat_list, second_army_combat_list = parser_obj.create_both_armies()

    first_army_with_shortage = shortage_before_and_after_battle(first_army_combat_list)
    second_army_with_shortage = shortage_before_and_after_battle(second_army_combat_list)

    first_army_first_stage, first_army_graveyard_1st_stage = first_stage_battle(second_army_with_shortage,
                                                                                first_army_with_shortage)
    second_army_first_stage, second_army_graveyard_1st_stage = first_stage_battle(first_army_with_shortage,
                                                                                  second_army_with_shortage)

    (first_army_second_stage, second_army_second_stage,
     first_army_graveyard_2nd_stage, second_army_graveyard_2nd_stage) = second_stage_battle(first_army_first_stage,
                                                                                            second_army_first_stage)

    first_army_final_graveyard = casualties_for_graveyard(first_army_graveyard_1st_stage +
                                                          first_army_graveyard_2nd_stage)
    second_army_final_graveyard = casualties_for_graveyard(second_army_graveyard_1st_stage +
                                                           second_army_graveyard_2nd_stage)

    first_army_with_casualties = shortage_before_and_after_battle(first_army_second_stage)
    second_army_with_casualties = shortage_before_and_after_battle(second_army_second_stage)

    file_writer(first_army_with_casualties, second_army_with_casualties,
                first_army_final_graveyard, second_army_final_graveyard)


if __name__ == "__main__":
    main_logic()
