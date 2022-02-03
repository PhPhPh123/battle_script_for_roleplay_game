import openpyxl as op
from random import shuffle, randint

"""
Global parameters using by functions
"""
first_pars_list = []  # lists of first and second army parsed from xlsx file, same units does not divide by quantity
second_pars_list = []

first_army_combat_list = []  # parsed lists where all same units divided by quantity
second_army_combat_list = []
first_army_magic_and_abilities_list = []
second_army_magic_and_abilities_list = []

morale_modifier = 2  # this mod multiplies combat points if unit will be lucky

base_casualties = 30


def first_stage_battle():
    for first_army_unit in first_army_combat_list:
        for second_army_unit in second_army_combat_list:
            if second_army_unit[0] in first_army_unit[6] and second_army_unit[2] > 0:
                dmg1 = second_army_unit[1] // 100 * (first_army_unit[3] - second_army_unit[4])
                if dmg1 > 0:
                    second_army_unit[2] -= dmg1
                shuffle(second_army_combat_list)
                break

    for second_army_unit in second_army_combat_list:
        for first_army_unit in first_army_combat_list:
            if first_army_unit[0] in second_army_unit[6] and first_army_unit[2] > 0:
                dmg2 = first_army_unit[1] // 100 * (second_army_unit[3] - first_army_unit[4])
                if dmg2 > 0:
                    first_army_unit[2] -= dmg2
                shuffle(first_army_combat_list)
                break


def combat_points_counter():
    first_army_sum_combat = 0  # final combat points of first and second army
    second_army_sum_combat = 0

    for unit in first_army_combat_list:
        if unit[2] > 0:
            if 0 < unit[5] > randint(0, 100):
                first_army_sum_combat += unit[2] * morale_modifier
            elif 0 > unit[5] >= randint(0, -100):
                first_army_sum_combat += unit[2] // morale_modifier
            else:
                first_army_sum_combat += unit[2]

    for unit in second_army_combat_list:
        if unit[2] > 0:
            if 0 < unit[5] > randint(0, 100):
                second_army_sum_combat += unit[2] * morale_modifier
            elif 0 > unit[5] >= randint(0, -100):
                second_army_sum_combat += unit[2] // morale_modifier
            else:
                second_army_sum_combat += unit[2]

    return first_army_sum_combat, second_army_sum_combat


def army_advantage(first_cb, second_cb):
    return int((first_cb - second_cb) / min(first_cb, second_cb) * 100 / 5)


def second_stage_battle():
    first_army_final_cb, second_army_final_cb = combat_points_counter()
    advantage = army_advantage(first_army_final_cb, second_army_final_cb)

    if advantage > 0:
        first_army_casualties = base_casualties - advantage // 2
        second_army_casualties = base_casualties + advantage
    elif advantage < 0:
        first_army_casualties = base_casualties + abs(advantage)
        second_army_casualties = base_casualties - abs(advantage // 2)
    else:
        first_army_casualties = base_casualties
        second_army_casualties = base_casualties

    if first_army_casualties >= 100:
        first_army_casualties = 100
    if first_army_casualties <= 0:
        first_army_casualties = 0
    if second_army_casualties >= 100:
        second_army_casualties = 100
    if second_army_casualties <= 0:
        second_army_casualties = 0

    for unit in first_army_combat_list:
        unit[2] = int(unit[2] // 100 * (100 - first_army_casualties))
    for unit in second_army_combat_list:
        unit[2] = int(unit[2] // 100 * (100 - second_army_casualties))

    return remove_dead_units(first_army_combat_list), remove_dead_units(second_army_combat_list)


def remove_dead_units(unitlist: list) -> list:
    new_list = []
    for unit in unitlist:
        if unit[2] > 0:
            new_list.append(unit)
    return new_list


def military_intelligence():
    pass


def magic():
    pass


def general_abilities():
    pass


def army_parser():
    xls_workbook = op.load_workbook("table.xlsx", read_only=True)
    xls_worksheet = xls_workbook.active

    for row in xls_worksheet.iter_rows(min_row=4, max_row=35, min_col=3, max_col=10, values_only=True):
        if not row[0]:
            break
        first_pars_list.append(list(row))
    for row in xls_worksheet.iter_rows(min_row=4, max_row=35, min_col=14, max_col=21, values_only=True):
        if not row[0]:
            break
        second_pars_list.append(list(row))

    for i in first_pars_list:
        while i[7] >= 1:
            first_army_combat_list.append(i[:-1])
            i[7] -= 1

    for i in second_pars_list:
        while i[7] >= 1:
            second_army_combat_list.append(i[:-1])
            i[7] -= 1

    for row in xls_worksheet.iter_rows(min_row=41, max_row=55, min_col=3, max_col=6, values_only=True):
        if not row[0]:
            break
        first_army_magic_and_abilities_list.append(list(row))

    for row in xls_worksheet.iter_rows(min_row=41, max_row=55, min_col=14, max_col=17, values_only=True):
        if not row[0]:
            break
        second_army_magic_and_abilities_list.append(list(row))

    print(first_army_magic_and_abilities_list)
    print(second_army_magic_and_abilities_list)

    shuffle(first_army_combat_list)
    shuffle(second_army_combat_list)


def magic_parser():
    pass


def file_writer(first_list, second_list):
    with open("result.txt", 'w', encoding="utf-8") as res:
        res.write("Первая армия выжившие:\n")
        for line in first_list:
            res.write(str(line) + '\n')
        res.write("Вторая армия выжившие:\n")
        for line in second_list:
            res.write(str(line) + '\n')


def main_logic():
    army_parser()
    first_stage_battle()
    second_stage_battle()
    first_army_afterfight, second_army_afterfight = second_stage_battle()
    file_writer(first_army_afterfight, second_army_afterfight)


if __name__ == "__main__":
    main_logic()
