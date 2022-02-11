from unittest import TestCase, main
import BattleScript.main as script


class ParserTest(TestCase):
    def test_length(self):
        self.assertEqual(len(script.parser()), 9)
        self.assertEqual(len(script.parser()), 9)
