from unittest import TestCase, main
import BattleScript.main as script
import openpyxl as op
import random
import pathlib

path = pathlib.Path(pathlib.Path.home(), "PycharmProjects", "battle_script_for_roleplay_game", 'BattleScript',
                    "table.xlsx")
xls_workbook = op.load_workbook(path, read_only=True)
xls_worksheet = xls_workbook.active


def unit_creator(stage):
    keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                          "damage", "defence", "morale", "targets", "amount")
    type_unit_list = ['в', "л", "к", "м", "ф", "ч", "п", "о"]

    unit_name = str("n" * random.randint(2, 6))
    unit_type = random.choice(type_unit_list)
    max_combat = random.randint(50, 500)
    current_combat = random.randint(1, max_combat)
    unit_damage = random.randint(0, 20)
    unit_defence = random.randint(0, 20)
    unit_morale = random.randint(0, 20)

    unit_targets = ""
    for _ in type_unit_list:
        random_target = random.choice(type_unit_list)
        if random_target not in unit_targets:
            unit_targets += f"{random_target}+"
        if len(unit_targets) > 4:
            break

    units_amount = random.randint(1, 5)

    values_for_dicts = (unit_name, unit_type, max_combat, current_combat, unit_damage, unit_defence,
                        unit_morale, unit_targets, units_amount)

    if stage == "parsing":
        return {keys: values for keys, values in zip(keys_for_army_dict, values_for_dicts)}
    elif stage != "parsing":
        return dict(list({keys: values for keys, values in zip(keys_for_army_dict, values_for_dicts)}.items())[:-1])


def army_creator(number_of_units, stage):
    army_list = []
    if stage == "parsing":
        positive_or_negative_combat = 1
    else:
        positive_or_negative_combat = random.choice((-1, 1))
    for _ in range(number_of_units):
        army_list.append(unit_creator(positive_or_negative_combat))
    return army_list


class ParserTest(TestCase):

    def test_number_of_units_after_parsing(self):
        self.assertNotEqual(script.army_list_creator(army_creator(random.randint(1, 10), "parsing")),
                                                    army_creator(random.randint(1, 10), "first-stage"))

    def test_army_without_unit_renaming(self):
        list_for_testing = [{"unit_name": "название_юнита", "unit_type": "в", "max_combat": 100, "current_combat": 75,
                             "damage": 15, "defence": 10, "morale": 10, "targets": "в,л,о", "amount": 2}]
        bad_result = [{"unit_name": "название_юнита", "unit_type": "в", "max_combat": 100, "current_combat": 75,
                       "damage": 15, "defence": 10, "morale": 10, "targets": "в,л,о"}] * 2
        self.assertNotEqual(script.army_list_creator(list_for_testing), bad_result)
