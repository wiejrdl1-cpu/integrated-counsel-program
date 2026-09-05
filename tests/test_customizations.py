import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import pastel_theme
import team1_app
import team2_app


class DummyEntry:
    def __init__(self, value="", state="normal"):
        self.value = value
        self.state = state

    def get(self):
        return self.value

    def delete(self, *_args):
        self.value = ""

    def insert(self, _index, value):
        self.value = value

    def config(self, **kwargs):
        self.state = kwargs.get("state", self.state)

    configure = config

    def cget(self, name):
        if name == "state":
            return self.state
        raise KeyError(name)


class DummyValue:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value


class HousingRuleTests(unittest.TestCase):
    def test_teams_share_the_same_housing_options(self):
        self.assertEqual(team1_app.HOME_OPTIONS, team2_app.HOME_OPTIONS)
        self.assertIn("민간(월세)", team1_app.HOME_OPTIONS)

    def assert_states(self, app_class, selector_name, housing_type, expected):
        app = object.__new__(app_class)
        app.e_home_direct = DummyEntry("old")
        app.e_deposit = DummyEntry("old")
        app.e_rent = DummyEntry("old")
        setattr(app, selector_name, DummyValue(housing_type))
        app_class.on_home_change(app)
        actual = (app.e_home_direct.state, app.e_deposit.state, app.e_rent.state)
        self.assertEqual(expected, actual, housing_type)

    def test_team1_housing_input_states(self):
        cases = {
            "자가": ("disabled", "disabled", "disabled"),
            "민간(전세)": ("disabled", "normal", "disabled"),
            "민간(보증부월세)": ("disabled", "normal", "normal"),
            "민간(월세)": ("disabled", "disabled", "normal"),
            "전세임대": ("disabled", "disabled", "normal"),
            "기타(직접입력)": ("normal", "normal", "normal"),
        }
        for housing_type, expected in cases.items():
            self.assert_states(team1_app.App, "home_var", housing_type, expected)

    def test_team2_housing_input_states_match_team1(self):
        for housing_type in team1_app.HOME_OPTIONS:
            app1 = object.__new__(team1_app.App)
            app1.home_var = DummyValue(housing_type)
            app1.e_home_direct = DummyEntry()
            app1.e_deposit = DummyEntry()
            app1.e_rent = DummyEntry()
            team1_app.App.on_home_change(app1)

            app2 = object.__new__(team2_app.App)
            app2.cb_home = DummyValue(housing_type)
            app2.e_home_direct = DummyEntry()
            app2.e_deposit = DummyEntry()
            app2.e_rent = DummyEntry()
            team2_app.App.on_home_change(app2)

            states1 = (app1.e_home_direct.state, app1.e_deposit.state, app1.e_rent.state)
            states2 = (app2.e_home_direct.state, app2.e_deposit.state, app2.e_rent.state)
            self.assertEqual(states1, states2, housing_type)

    def test_disabled_entry_color_is_defined(self):
        self.assertEqual("#E2E6E3", pastel_theme.PALETTE["entry_disabled"])

    def test_team1_supporter_has_no_income_field(self):
        self.assertNotIn("e_income", team1_app.SupporterRow.values.__code__.co_names)
        self.assertNotIn("income", team1_app.SupporterRow.values.__code__.co_consts)


if __name__ == "__main__":
    unittest.main()
