from unittest import TestCase, main
import BattleScript.main as script
import openpyxl as op
import random
import pathlib

path = pathlib.Path(pathlib.Path.home(), "PycharmProjects", "battle_script_for_roleplay_game", 'BattleScript',
                    "table.xlsx")
xls_workbook = op.load_workbook(path, read_only=True)
xls_worksheet = xls_workbook.active


def list_army_testing_creator(positive_or_negative):
    keys_for_army_dict = ("unit_name", "unit_type", "max_combat", "current_combat",
                          "damage", "defence", "morale", "targets", "amount")

    unit_name_ = str("n" * random.randint(2, 6))
    unit_type = str("n" * random.randint(9, 15))
    max_combat = random.randint(50, 500)
    current_combat = random.randint(1, max_combat) * positive_or_negative
    unit_damage = random.randint(0, 20)
    unit_defence = random.randint(0, 20)
    unit_morale = random.randint(0, 20)

    values_for_dicts = ()


class ParserTest(TestCase):

    def test_army_without_unit_renaming(self):
        list_for_testing = [{"unit_name": "название_юнита", "unit_type": "в", "max_combat": 100, "current_combat": 75,
                             "damage": 15, "defence": 10, "morale": 10, "targets": "в,л,о", "amount": 2}]
        bad_result = [{"unit_name": "название_юнита", "unit_type": "в", "max_combat": 100, "current_combat": 75,
                       "damage": 15, "defence": 10, "morale": 10, "targets": "в,л,о"}] * 2
        self.assertNotEqual(script.army_list_creator(list_for_testing), bad_result)
