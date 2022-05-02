import openpyxl as op
from random import shuffle, randint
from pprint import pprint

morale_modifier = 2  # this mod multiplies combat points if unit will be lucky. Used in second battle stage
base_casualties = 30  # base level of casualties(in fact - percents) used in second battle stage

winner = None
second_stage_advantage = 0


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

                    break  # this break in fact return us to first "for" cycle. Its need to prevent hitting more then 1 unit

    defender_after_casualties = remove_dead_units(list.copy(defender_list))

    return defender_after_casualties


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
                army_sum_combat += unit["current_combat"] * morale_modifier  # add combat point * morale to sum
            elif 0 > unit["morale"] >= randint(-100, 0):  # negative morale check(morale<0).
                # check succeeds if morale >= randint
                army_sum_combat += int(unit["current_combat"] / morale_modifier)  # add combat point // morale to sum
            else:
                army_sum_combat += unit["current_combat"]  # add combat points to sum if morale == 0

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

    first_army_afterfight = remove_dead_units(first_army)
    second_army_afterfight = remove_dead_units(second_army)

    return first_army_afterfight, second_army_afterfight


def remove_dead_units(unitlist: list) -> list:
    new_list = []
    for unit in unitlist:
        if unit["current_combat"] > 0:
            new_list.append(unit)
    return new_list


def parser(worksheet, keys_dict, minrow, maxrow, mincol, maxcol):
    pars_list = []

    for row in worksheet.iter_rows(min_row=minrow, max_row=maxrow, min_col=mincol, max_col=maxcol,
                                   values_only=True):
        if not row[0]:
            break
        army_or_ability_dict = dict(zip(keys_dict, row))
        pars_list.append(army_or_ability_dict)

    return pars_list


def army_unit_numeration(arm_list):
    sorted_arm_list = sorted(arm_list, key=lambda islist: islist["unit_name"])

    for num, unit in enumerate(sorted_arm_list, start=1):
        unit["unit_name"] += f" ,юнит №{num}"
    numerated_arm_list = sorted_arm_list.copy()

    return numerated_arm_list


def army_list_creator(parslist):
    army_list = []

    for i in parslist:
        while i["amount"] >= 1:
            army_list.append(dict(list(i.items())[:-1]))
            i["amount"] -= 1
    numerated_army_list = army_unit_numeration(army_list)

    return numerated_army_list


def pop_casualties(army_list):
    casualties = 0
    for unit in army_list:
        casualties += int(100 - (unit["current_combat"] / unit['max_combat'] * 100))
    return casualties


def final_casualites_in_pops(casualty_before_battle: int, sum_of_casualty_after_battle) -> int:
    return sum_of_casualty_after_battle - casualty_before_battle


def file_writer(first_list, second_list, first_army_pop_cas, second_army_pop_cas):
    with open("result.txt", 'w', encoding="utf-8") as res:
        res.write(f'Бой проведен. Его результатом стала {winner}')
        res.write(f'\nПреемущество победителя во второй фазе боя - {abs(second_stage_advantage)}%\n')

        if first_list:
            res.write("\nПервая армия выжившие:\n")
            for line in first_list:
                res.write(str(line["unit_name"] + ' оставшаяся комбатка: ' + str(line['current_combat']) + '\n'))
        else:
            res.write("Первая армия полностью уничтожена")

        if second_list:
            res.write("\nВторая армия выжившие:\n")
            for line in second_list:
                res.write(str(line["unit_name"]) + ' оставшаяся комбатка: ' + str(line['current_combat']) + '\n')
        else:
            res.write("\nВторая армия полностью уничтожена")

        res.write(f"\nПервая армия потеряла {first_army_pop_cas} юнитов(населения)\n")
        res.write(f"Вторая армия потеряла {second_army_pop_cas} юнитов(населения)")


def main_logic():
    xls_workbook = op.load_workbook("table.xlsx", read_only=True)
    xls_worksheet = xls_workbook.active

    keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                          "damage", "defence", "morale", "targets", "combat_bonus", "amount")

    first_army_pars_list = parser(xls_worksheet, keys_for_army_dict, minrow=4, maxrow=35, mincol=2, maxcol=11)
    second_army_pars_list = parser(xls_worksheet, keys_for_army_dict, minrow=4, maxrow=35, mincol=14, maxcol=23)

    first_army_combat_list = army_list_creator(first_army_pars_list)
    second_army_combat_list = army_list_creator(second_army_pars_list)

    first_army_shortage_before_the_battle = pop_casualties(first_army_combat_list)
    second_army_shortage_before_the_battle = pop_casualties(second_army_combat_list)

    first_army_first_stage = first_stage_battle(second_army_combat_list, first_army_combat_list)
    second_army_first_stage = first_stage_battle(first_army_combat_list, second_army_combat_list)

    first_army_second_stage, second_army_second_stage = second_stage_battle(first_army_first_stage,
                                                                            second_army_first_stage)

    first_army_pop_cas = pop_casualties(first_army_second_stage)
    second_army_pop_cas = pop_casualties(second_army_second_stage)

    first_army_sum_of_casualies = final_casualites_in_pops(first_army_shortage_before_the_battle, first_army_pop_cas)
    second_army_sum_of_casualies = final_casualites_in_pops(second_army_shortage_before_the_battle, second_army_pop_cas)

    file_writer(first_army_second_stage, second_army_second_stage,
                first_army_sum_of_casualies, second_army_sum_of_casualies)


if __name__ == "__main__":
    main_logic()
