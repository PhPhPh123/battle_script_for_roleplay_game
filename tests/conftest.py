import pytest
import openpyxl as op
from copy import deepcopy
import random
from string import ascii_lowercase


@pytest.fixture()
def default_armies_fixture():
    xls_workbook = op.load_workbook("table.xlsx", read_only=True)
    xls_worksheet = xls_workbook.active
    keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                          "damage", "defence", "morale", "targets", "combat_bonus", "amount")
    raws_and_columns_first_army = (4, 35, 2, 11)
    raws_and_columns_second_army = (4, 35, 14, 23)

    def parse_raw(min_row, max_row, min_col, max_col):
        pars_list = []  # инициализация пустого листа который впоследствии станет списком словарей
        for row in xls_worksheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col,
                                           values_only=True):
            if not row[0]:  # если парсер доходит до пустой строки, значит юниты кончились и парсинг останавливается
                break
            army_or_ability_dict = dict(zip(keys_for_army_dict, row))  # отпаршенные значения связываются с ключами
            pars_list.append(army_or_ability_dict)  # словари добавляются в общий список

        return pars_list

    def create_army_dicts(parslist: list[dict, ...]) -> list[dict, ...]:
        army_list = []  # инициализация списка словарей
        for i in parslist:
            # создание отдельных юнитов продолжается только когда их количество в стаке больше или равно 1, если
            # при достижении 0 разбитие стака на отдельные юниты прекращается
            while i["amount"] >= 1:
                # в список отдельных юнитов добавляются словари со всеми ключ-значениями кроме последнего,
                # т.к. он отвечает за количество юнитов в стаке [:-1]
                army_list.append(dict(list(i.items())[:-1]))
                i["amount"] -= 1  # уменьшаю значение стака на 1
        return army_list

    def numerate_army_units(arm_list: list[dict, ...]) -> list[dict, ...]:
        # список словарей сортируется по именам юнитов
        sorted_arm_list = sorted(arm_list, key=lambda islist: islist["unit_name"])

        # нумерация отсортированного списка и придание каждому юниту его номера
        for num, unit in enumerate(sorted_arm_list, start=1):
            unit["unit_name"] += f" ,юнит №{num}"
        numerated_arm_list = deepcopy(sorted_arm_list)  # копирую список глубоким копированием
        return numerated_arm_list

    first_army_parse = parse_raw(*raws_and_columns_first_army)
    second_army_parse = parse_raw(*raws_and_columns_second_army)

    first_army_create = create_army_dicts(first_army_parse)
    second_army_create = create_army_dicts(second_army_parse)

    first_army_numerated = numerate_army_units(first_army_create)
    second_army_numerated = numerate_army_units(second_army_create)

    return first_army_numerated, second_army_numerated


@pytest.fixture()
def random_units_stack_fixture():
    stack_amount = random.randint(1, 5)
    units_stack = create_unit(stack=True, stack_amount=stack_amount)
    return units_stack


@pytest.fixture()
def random_single_unit_fixture():
    unit = create_unit()
    return unit


@pytest.fixture()
def random_army_creator_with_stacks_fixture(random_units_stack_fixture):
    army_list = []
    random_units_amount = random.randint(1, 20)

    for _ in range(random_units_amount):
        army_list.append(random_units_stack_fixture)
    return army_list


@pytest.fixture()
def random_army_creator_with_single_units_fixture(random_single_unit_fixture):
    army_list = []
    random_units_amount = random.randint(1, 20)

    for _ in range(random_units_amount):
        army_list.append(random_single_unit_fixture)
    return army_list


def create_unit(stack=False, stack_amount=None):
    keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                          "damage", "defence", "morale", "targets", "amount")
    type_unit_list = ['в', "л", "к", "м", "ф", "ч", "п", "о"]

    unit_name = "".join([random.choice(ascii_lowercase) for _ in range(10)])
    unit_type = random.choice(type_unit_list)
    max_combat = random.randint(50, 500)
    current_combat = random.randint(1, max_combat)
    unit_damage = random.randint(0, 20)
    unit_defence = random.randint(0, 20)
    unit_morale = random.randint(0, 20)
    units_amount = stack_amount

    unit_targets = ""
    for _ in type_unit_list:
        random_target = random.choice(type_unit_list)
        if random_target not in unit_targets:
            unit_targets += f"{random_target}+"
        if len(unit_targets) > 4:
            break

    if stack:
        values_for_dicts = (unit_name, unit_type, max_combat, current_combat, unit_damage, unit_defence,
                            unit_morale, unit_targets, units_amount)
    else:
        values_for_dicts = (unit_name, unit_type, max_combat, current_combat, unit_damage, unit_defence,
                            unit_morale, unit_targets)

    if stack:
        return dict(list({keys: values for keys, values in zip(keys_for_army_dict, values_for_dicts)}.items())[:-1])
    else:
        return {keys: values for keys, values in zip(keys_for_army_dict, values_for_dicts)}
