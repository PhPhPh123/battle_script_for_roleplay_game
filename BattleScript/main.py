import openpyxl as op
from random import shuffle, randint

"""
lists of first and second army parsed from xlsx file, same units does not divide by quantity
***_pars_list[7] its amount of units with the same parameters. In fact - stack of units
"""
first_pars_list = []
second_pars_list = []

"""
В этих списках отсутвует параметр количества и стаки юнитов разбиты на отдельных юнитов
In this lists missing amount parameter and stacks of units divide by individual unit
element in this lists its list of unit parameters:
***_army_combat_list[*][0] its type of unit displayed by one string sign in russian language
***_army_combat_list[*][1] its maximum combat points of unit
***_army_combat_list[*][2] its current combat points of unit
***_army_combat_list[*][3] its damage of unit in first stage
***_army_combat_list[*][4] its unit damage protection
***_army_combat_list[*][5] its moral parameter which doubled or halved combat points in second battle stage
***_army_combat_list[*][6] its string of russian signs which show what type of troops will receive damage
"""
first_army_combat_list = []  # parsed lists where all same units divided by quantity
second_army_combat_list = []
first_army_magic_and_abilities_list = []
second_army_magic_and_abilities_list = []

morale_modifier = 2  # this mod multiplies combat points if unit will be lucky. Used in second battle stage
base_casualties = 30  # base level of casualties(in fact - percents) used in second battle stage


def first_stage_battle() -> None:
    """
    this function counts first stage battle and modified army lists to use it in second battle stage
    :return: None
    """

    """
    this cycle applies magic in the first stage of the battle
    the first "if" operator selects the type of magic to be used
    """
    for spell_and_ability in first_army_magic_and_abilities_list:
        if spell_and_ability[0] == "а":
            for second_army_unit in second_army_combat_list:
                second_army_unit[2] -= spell_and_ability[1]
                spell_and_ability[2] -= 1
                if spell_and_ability[2] == 0:
                    break

    for first_army_unit in first_army_combat_list:  # iterate by unit in first army list
        for second_army_unit in second_army_combat_list:  # iterate by unit in second army list
            if second_army_unit[0] in first_army_unit[6] and second_army_unit[2] > 0:  # if unit in second army have
                # type which damaged by first army unit and second army have at least 1 combat point(HP)
                dmg1 = second_army_unit[1] // 100 * (first_army_unit[3] - second_army_unit[4])  # this formula counts
                # damage
                if dmg1 > 0:  # if protection bigger then the damage then no damage is dealt
                    second_army_unit[2] -= dmg1  # dealt damage
                shuffle(second_army_combat_list)  # random function which shuffle army list in each iteration
                #  This is necessary to introduce an element of chance
                break  # this break in fact return us to first "for" cycle. Its need to prevent hitting more then 1 unit

    for second_army_unit in second_army_combat_list:
        for first_army_unit in first_army_combat_list:
            if first_army_unit[0] in second_army_unit[6] and first_army_unit[2] > 0:
                dmg2 = first_army_unit[1] // 100 * (second_army_unit[3] - first_army_unit[4])
                if dmg2 > 0:
                    first_army_unit[2] -= dmg2
                shuffle(first_army_combat_list)
                break


def combat_points_counter() -> (int, int):
    """
    this function counts sum of combat points to use it in second battle stage
    also this function use moral modifier to double (or halve) combat points if unit is lucky
    :return: sum of first and second army combat points
    """
    first_army_sum_combat = 0  # final combat points of first and second army
    second_army_sum_combat = 0

    for unit in first_army_combat_list:  # iterate by unit in army list
        if unit[2] > 0:  # if unit is not dead(combat point <= 0)
            if 0 < unit[5] >= randint(0, 100):  # positive morale check(morale>0).check succeeds if morale >= randint
                first_army_sum_combat += unit[2] * morale_modifier  # add combat point * morale to sum
            elif 0 > unit[5] >= randint(0, -100):  # negative morale check(morale<0).check succeeds if morale >= randint
                first_army_sum_combat += unit[2] // morale_modifier  # add combat point // morale to sum
            else:
                first_army_sum_combat += unit[2]  # add combat points to sum if morale == 0

    for unit in second_army_combat_list:
        if unit[2] > 0:
            if 0 < unit[5] > randint(0, 100):
                second_army_sum_combat += unit[2] * morale_modifier
            elif 0 > unit[5] >= randint(0, -100):
                second_army_sum_combat += unit[2] // morale_modifier
            else:
                second_army_sum_combat += unit[2]

    return first_army_sum_combat, second_army_sum_combat


def army_advantage(first_cb: int, second_cb: int) -> int:
    """
    this function counts advantage of army with bigger combat points expressed as a percentage
    :param first_cb: first army combat point
    :param second_cb: second army combat points
    :return: advantage of the first army over the second, if advantage < 0 it means first army is weaker them second
    """
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


def magic_and_abilities():
    for spell_and_ability in first_army_magic_and_abilities_list:
        if spell_and_ability[0] == "а":
            return


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

    # print(first_army_magic_and_abilities_list)
    # print(second_army_magic_and_abilities_list)

    shuffle(first_army_combat_list)
    shuffle(second_army_combat_list)


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
