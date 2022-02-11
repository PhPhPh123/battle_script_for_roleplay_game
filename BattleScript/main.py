import openpyxl as op
from random import shuffle, randint

"""
lists of first and second army parsed from xlsx file, same units does not divide by quantity
***_pars_list[7] its amount of units with the same parameters. In fact - stack of units
"""

"""
В этих списках отсутвует параметр количества и стаки юнитов разбиты на отдельных юнитов
In this lists missing amount parameter and stacks of units divide by individual unit
element in this lists its list of unit parameters:
***_army_combat_list[*][0] its name of unit
***_army_combat_list[*][1] its type of unit displayed by one string sign in russian language
***_army_combat_list[*][2] its maximum combat points of unit
***_army_combat_list[*][3] its current combat points of unit
***_army_combat_list[*][4] its damage of unit in first stage
***_army_combat_list[*][5] its unit damage protection
***_army_combat_list[*][6] its moral parameter which doubled or halved combat points in second battle stage
***_army_combat_list[*][7] its string of russian signs which show what type of troops will receive damage
"""

morale_modifier = 2  # this mod multiplies combat points if unit will be lucky. Used in second battle stage
base_casualties = 30  # base level of casualties(in fact - percents) used in second battle stage

winner = None
second_stage_advantage = 0


def first_stage_battle(attacker_list, defender_list):
    """
    this function counts first stage battle and modified army lists to use it in second battle stage
    :return: None
    """
    count = 0
    for attacker_army_unit in attacker_list:  # iterate by unit in first army list
        for defender_army_unit in defender_list:  # iterate by unit in second army list

            # if unit in second army have type which damaged by first army unit and second army have
            # at least 1 combat point(HP)
            if defender_army_unit["unit_type"] in attacker_army_unit["targets"] and defender_army_unit["current_combat"] > 0:

                # this formula counts damage
                dmg1 = defender_army_unit["max_combat"] // 100 * (attacker_army_unit["damage"] - defender_army_unit["defence"])

                if dmg1 > 0:  # if protection bigger then the damage then no damage is dealt(it needs to prevent
                    # negative damage)
                    defender_army_unit["current_combat"] -= dmg1  # dealt damage to unit
                    count += 1


                shuffle(defender_list)  # random function which shuffle army list in each iteration
                #  This is necessary to introduce an element of chance

                break  # this break in fact return us to first "for" cycle. Its need to prevent hitting more then 1 unit
    print(count)
    defender_after_casualties = list.copy(defender_list)
    for i in defender_after_casualties:
        print(i["current_combat"])

    return defender_after_casualties


def combat_points_counter() -> (int, int):
    """
    this function counts sum of combat points to use it in second battle stage
    also this function use moral modifier to double (or halve) combat points if unit is lucky
    :return: sum of first and second army combat points
    """
    first_army_sum_combat = 0  # final combat points of first and second army
    second_army_sum_combat = 0

    for unit in first_army_combat_list:  # iterate by unit in army list
        if unit["current_combat"] > 0:  # if unit is not dead(combat point <= 0)
            if 0 < unit["morale"] >= randint(0, 100):  # positive morale check(morale>0).
                # check succeeds if morale >= randint
                first_army_sum_combat += unit["current_combat"] * morale_modifier  # add combat point * morale to sum
            elif 0 > unit["morale"] >= randint(0, -100):  # negative morale check(morale<0).
                # check succeeds if morale >= randint
                first_army_sum_combat += unit["current_combat"] // morale_modifier  # add combat point // morale to sum
            else:
                first_army_sum_combat += unit["current_combat"]  # add combat points to sum if morale == 0

    for unit in second_army_combat_list:
        if unit["current_combat"] > 0:
            if 0 < unit["morale"] > randint(0, 100):
                second_army_sum_combat += unit["current_combat"] * morale_modifier
            elif 0 > unit["morale"] >= randint(0, -100):
                second_army_sum_combat += unit["current_combat"] // morale_modifier
            else:
                second_army_sum_combat += unit["current_combat"]

    global winner
    if first_army_sum_combat > second_army_sum_combat:
        winner = "победа первой армии"
    elif first_army_sum_combat < second_army_sum_combat:
        winner = "победа второй армии"
    else:
        winner = "Ничья"
    return first_army_sum_combat, second_army_sum_combat


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


def second_stage_battle():
    first_army_final_cb, second_army_final_cb = combat_points_counter()
    if first_army_final_cb == 0 or second_army_final_cb == 0:
        return remove_dead_units(first_army_combat_list), remove_dead_units(second_army_combat_list)
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
        unit["current_combat"] = int(100 / unit["current_combat"] * (100 - first_army_casualties))
    for unit in second_army_combat_list:
        print(unit)
        unit["current_combat"] = int(100 / unit["current_combat"] * (100 - first_army_casualties))
    return remove_dead_units(first_army_combat_list), remove_dead_units(second_army_combat_list)


def remove_dead_units(unitlist: list) -> list:
    new_list = []
    for unit in unitlist:
        if unit["current_combat"] > 0:
            new_list.append(unit)
    return new_list


def military_intelligence():
    pass


def magic_and_abilities():
    """
    this cycle applies magic in the first stage of the battle
    the first "if" operator selects the type of magic to be used
    """
    for spell_and_ability in first_army_magic_and_abilities_list:
        if spell_and_ability['name_of_magic_or_abilities'] == "а":
            for second_army_unit in second_army_combat_list:
                second_army_unit["current_combat"] -= spell_and_ability['skill_damage']
                spell_and_ability['number_of_targets'] -= 1
                if spell_and_ability['number_of_targets'] == 0:
                    break


def parser(worksheet, keys_dict, minrow, maxrow, mincol, maxcol):
    pars_list = []

    for row in worksheet.iter_rows(min_row=minrow, max_row=maxrow, min_col=mincol, max_col=maxcol,
                                   values_only=True):
        if not row[0]:
            break
        army_or_ability_dict = dict(zip(keys_dict, row))
        pars_list.append(army_or_ability_dict)

    return pars_list


def army_list_creator(parslist):
    army_list = []

    for i in parslist:
        while i["amount"] >= 1:
            army_list.append(i)
            i["amount"] -= 1
    return army_list


def file_writer(first_list, second_list):
    with open("result.txt", 'w', encoding="utf-8") as res:
        res.write(f'Бой проведен. Его результатом стала {winner}')
        res.write(f'\nПреемущество победителя во второй фазе боя - {abs(second_stage_advantage)}%\n')

        if first_list:
            res.write("\nПервая армия выжившие:\n")
            for line in first_list:
                res.write(str(line["unit_name"] + str(line['current_combat']) + '\n'))
        else:
            res.write("Первая армия полностью уничтожена")

        if second_list:
            res.write("\nВторая армия выжившие:\n")
            for line in second_list:
                res.write(str(line["unit_name"]) + ' ' + str(line['current_combat']) + '\n')
        else:
            res.write("\nВторая армия полностью уничтожена")


def main_logic():
    xls_workbook = op.load_workbook("table.xlsx", read_only=True)
    xls_worksheet = xls_workbook.active

    keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                          "damage", "defence", "morale", "targets", "amount")
    keys_for_magic_dict = ('name_of_magic_or_abilities', 'type_of_magic_or_abilities',
                           'skill_damage', 'number_of_targets', 'effectiveness')

    first_army_pars_list = parser(xls_worksheet, keys_for_army_dict, minrow=4, maxrow=35, mincol=2, maxcol=10)
    second_army_pars_list = parser(xls_worksheet, keys_for_army_dict, minrow=4, maxrow=35, mincol=13, maxcol=21)

    first_ability_list = parser(xls_worksheet, keys_for_magic_dict, minrow=41, maxrow=55, mincol=2, maxcol=6)
    second_ability_list = parser(xls_worksheet, keys_for_magic_dict, minrow=41, maxrow=55, mincol=13, maxcol=17)

    first_army_combat_list = army_list_creator(first_army_pars_list)
    second_army_combat_list = army_list_creator(second_army_pars_list)
    print(second_army_combat_list)

    first_army_first_stage = first_stage_battle(attacker_list=second_army_combat_list,
                                                defender_list=first_army_combat_list)
    print("")
    second_army_first_stage = first_stage_battle(attacker_list=first_army_combat_list,
                                                 defender_list=second_army_combat_list)



    first_army_afterfight, second_army_afterfight = second_stage_battle()
    file_writer(first_army_afterfight, second_army_afterfight)


if __name__ == "__main__":
    main_logic()
