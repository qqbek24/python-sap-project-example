# Author: Mateusz.Czub@arcelormittal.com

from pathlib import Path
from rpa_lib.template.lib_paths import Lib_Paths
from rpa_bot.bot_helpers import BotMode


class Paths(Lib_Paths):

    def __init__(self, robot, main_path: Path):
        log = robot.paths.log
        super().__init__(robot, main_path)
        self.log = log
        self.parameters = Path(f"{main_path}\\config.yaml")
        self.prod_folder = Path(f"\\\\vdi-your-path.com\\PL\\path\\{robot.robot_name}")
        self.test_folder = Path(f"\\\\vdi-your-path.com\\PL\\path-Test\\{robot.robot_name}")
        if robot.mode == BotMode.PROD:
            self.config = Path(f"{self.prod_folder}\\config.ini")
        elif robot.mode == BotMode.TEST:
            self.config = Path(f"{self.test_folder}\\config.ini")
