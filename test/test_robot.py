from robot import Robot
import pytest
import os
from pathlib import Path
from rpa_lib.robot_status import Robot_mode
from rpa_lib.log import lte, log
from rpa_lib.desktop.process import taskKill

# # For debug console issue while launching GUI apps
# import faulthandler 
# faulthandler.disable()

pytest.robot = None

@pytest.fixture
def robot():
    from robot import Robot
    abspath = os.path.dirname(os.path.abspath(__file__))
    main_path = Path(abspath).parent
    if pytest.robot is None:
        robot = Robot(main_path, "0.0.0.1", Robot_mode.dev)
        pytest.robot = robot
    return pytest.robot

def test_example1(robot: Robot):
    log("test_example1")

def error():
    raise Exception("test_example2")
def test_example2(robot: Robot):
    with pytest.raises(Exception) as exc_info:
        error()
    assert exc_info.value.args[0] == "test_example2"

# # For SAP robots
# from sap import SAP
# pytest.sap = None
# @pytest.fixture
# def sap(robot: Robot):
#     if pytest.sap is None:
#         pytest.sap = SAP(sap_sid, self.credentials.kp, self.credentials.kp_entry)
#         pytest.sap.connect()
#     return pytest.sap
# def test_example3(sap: SAP, robot: Robot):
#     pass

def test_clear():
    """Ending of test - close all windows opened by prevoius tests"""
    taskKill("saplogon.exe")
