import os
import time
import winreg
import shutil
import sys
import subprocess
import configparser
import json


class Setup:
    def __init__(self):
        self.bat_start = None
        self.bat_final = None
        self.run = 0
        self.close = 6
        self.close_target = 1
        self.config = configparser.ConfigParser()
        self.config.read('misc\\config.ini')

    @staticmethod
    def show_ascii():
        a = """
 ██████╗ ███╗   ███╗██████╗
 ██╔════╝████╗ ████║██╔══██╗
 ██║     ██╔████╔██║██║  ██║
 ██║     ██║╚██╔╝██║██║  ██║
 ╚██████╗██║ ╚═╝ ██║██████╔╝
 ╚═════╝╚═╝     ╚═╝╚═════╝"""
        return print(f'\n{a}')

    @staticmethod
    def check_registry_value_exists():
        """ checks to see if the expected registry value exists or not. """

        key = r"SOFTWARE\Microsoft\Command Processor"
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as reg_key:
                try:
                    value, _ = winreg.QueryValueEx(reg_key, "AutoRun")
                    return True
                except FileNotFoundError:
                    return False
        except FileNotFoundError:
            return False

    def check_and_update_ini(self) -> bool:
        """ Checks the config file to see if it's the first set up or not. """

        if self.config['ini']['first_run'] == '0':
            self.config['ini']['first_run'] = '1'
            with open('misc\\config.ini', 'w') as configfile:
                self.config.write(configfile)
            return True
        return False

    def config_bat(self) -> None:
        """ creates the .bat file to run every time cmd is opened. """

        # Creates the instance of the .bat file
        cwd = os.getcwd()
        with open('cmd.bat', 'w') as f:
            f.write('@echo off' + '\n')
            f.write(f'python "{cwd}\\misc\\main.py"')
            self.bat_start = f'{cwd}\\cmd.bat'

        # Saves the created .bat file to %localappdata%
        self.bat_final = f'{os.environ["LOCALAPPDATA"]}\\ASCLII'

        try:
            os.mkdir(self.bat_final)
            try:
                shutil.move(self.bat_start, self.bat_final)
            except shutil.Error:
                """ cmd.bat file already exists in CMDebug directory """
        except FileExistsError:
            """ CMDebug directory already exists in %localappdata% """
            os.remove(self.bat_start)

    def add_reg_key(self) -> None:
        """Adds the registry key to launch .bat every time CMD is opened."""

        key = r"SOFTWARE\Microsoft\Command Processor"
        name = "AutoRun"
        data = f'{self.bat_final}\\cmd.bat'

        try:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key, 0, winreg.KEY_WRITE)
            except FileNotFoundError:
                key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key)

            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, data)
            winreg.CloseKey(key)
            if self.check_and_update_ini():
                self.confirm_setup()

        except PermissionError:
            if not self.check_registry_value_exists():
                if not self.check_and_update_ini():
                    print('Error: Failed to add registry key due to no administrator permission.')
                    sys.exit(1)
                else:
                    self.confirm_setup()

    def confirm_setup(self):
        print(f'Setup complete. This window will close in {self.close -1} seconds.')

        while self.close != self.close_target:
            self.close -= 1
            print(f'{self.close}...')
            time.sleep(1)

        try:
            subprocess.run(['taskkill', '/IM', 'cmd.exe', '/F'], check=True)
        except subprocess.CalledProcessError:
            pass


class ChangeColour:
    def __init__(self):
        self.key = r"SOFTWARE\Microsoft\Command Processor"
        self.name = "DefaultColor"
        self.data = 0

    @staticmethod
    def set_colour():
        key = r"SOFTWARE\Microsoft\Command Processor"
        name = "DefaultColor"

        with open('misc\\colours.json', 'r') as json_file:
            data = json.load(json_file)

            for colour in data.keys():
                print(colour)
            colour_selected = input('Choose from the available colours: ').capitalize()
            color_value = data.get(colour_selected, '')

        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key, 0, winreg.KEY_WRITE)
        except FileNotFoundError:
            key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key)
        except PermissionError:
            print('Error: Failed to add registry key due to no administrator permission.')

        if not isinstance(color_value, str):
            color_value = str(color_value)

        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, color_value)
        winreg.CloseKey(key)


if '__main__' == __name__:
    setup = Setup()
    color = ChangeColour()

    if not setup.check_registry_value_exists():
        colour_decision = input('Would you like to select the color of your command line? [Y/N]: ').upper()
        if colour_decision == 'Y':
            color.set_colour()

        setup.config_bat()
        setup.add_reg_key()

    if setup.check_registry_value_exists():
        setup.show_ascii()
    else:
        print('Error: An error occurred, Most likely related to lack of Administrator Privileges.')
