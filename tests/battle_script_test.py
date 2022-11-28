import BattleScript.main as script
import openpyxl as op
import random
import pathlib


path = pathlib.Path(pathlib.Path.home(), "PycharmProjects", "battle_script_for_roleplay_game", 'BattleScript',
                    "default_army.xlsx")
xls_workbook = op.load_workbook(path, read_only=True)
xls_worksheet = xls_workbook.active


class ParserTest(TestCase):
    def test_type(self):
        bbb = army_creator(10, "parsing")
        ttt = script.army_list_creator(bbb)


    def test_number_of_units_after_parsing(self):
        self.assertNotEqual(script.army_list_creator(army_creator(random.randint(1, 10), "parsing")),
                            army_creator(random.randint(1, 10), "first-stage"))

    def test_army_without_unit_renaming(self):
        list_for_testing = [{"unit_name": "название_юнита", "unit_type": "в", "max_combat": 100, "current_combat": 75,
                             "damage": 15, "defence": 10, "morale": 10, "targets": "в,л,о", "amount": 2}]
        bad_result = [{"unit_name": "название_юнита", "unit_type": "в", "max_combat": 100, "current_combat": 75,
                       "damage": 15, "defence": 10, "morale": 10, "targets": "в,л,о"}] * 2
        self.assertNotEqual(script.army_list_creator(list_for_testing), bad_result)
