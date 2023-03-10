import datetime
import os
import warnings
from dataclasses import dataclass
from time import sleep
from typing import List
import dotenv
import psutil
import pywinauto
import requests
from pywinauto.application import ProcessNotFoundError
from pywinauto.controls.hwndwrapper import DialogWrapper
from pywinauto.findbestmatch import MatchError
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
from pywinauto.timings import TimeoutError as TimingsTimeoutError
from bot_notification import TelegramNotifier


@dataclass
class Credentials:
    usr: str
    psw: str


class BackendManager:
    def __init__(self, app: pywinauto.Application, backend_name: str) -> None:
        self.app, self.backend_name = app, backend_name

    def __enter__(self) -> None:
        self.app.backend.name = self.backend_name

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.app.backend.name = 'win32' if self.backend_name == 'uia' else 'uia'


class Utils:
    @staticmethod
    def get_window(title: str, app: pywinauto.Application, wait_for: str = 'exists',
                   timeout: int = 20, regex: bool = False) -> pywinauto.WindowSpecification:
        _window = app.window(title=title) if not regex else app.window(title_re=title)
        _window.wait(wait_for=wait_for, timeout=timeout)
        return _window

    @staticmethod
    def get_current_process_pid(proc_name: str) -> int or None:
        return next((p.pid for p in psutil.process_iter() if proc_name in p.name()), None)

    @staticmethod
    def kill_process(pid: int) -> None:
        p = psutil.Process(pid)
        p.terminate()

    @staticmethod
    def kill_colvirs() -> None:
        for proc in psutil.process_iter():
            try:
                if 'COLVIR' in proc.name():
                    process = psutil.Process(proc.pid)
                    process.terminate()
            except psutil.AccessDenied:
                continue


class Colvir:
    def __init__(self, credentials: Credentials, tomorrow: str) -> None:
        self.credentials = credentials
        self.pid: int or None = None
        self.app: pywinauto.Application or None = None
        self.tomorrow: str = tomorrow
        self.utils = Utils()
        self.utils.kill_colvirs()
        self.is_next_day: bool = False

    def login(self):
        desktop = pywinauto.Desktop(backend='win32')
        try:
            login_win = desktop.window(title='???????? ?? ??????????????')
            login_win.wait(wait_for='exists', timeout=20)
            login_win['Edit2'].wrapper_object().set_text(text=self.credentials.usr)
            login_win['Edit'].wrapper_object().set_text(text=self.credentials.psw)
            login_win['OK'].wrapper_object().click()
        except ElementAmbiguousError:
            windows: List[DialogWrapper] = pywinauto.Desktop(backend='win32').windows()
            for win in windows:
                if '???????? ?? ??????????????' not in win.window_text():
                    continue
                self.utils.kill_process(pid=win.process_id())
            raise ElementNotFoundError

    def confirm_warning(self) -> None:
        for window in self.app.windows():
            if window.window_text() != 'Colvir Banking System':
                continue
            win = self.app.window(handle=window.handle)
            for child in win.descendants():
                if child.window_text() == 'OK':
                    child.send_keystrokes('{ENTER}')
                    break

    def check_is_next_day(self) -> bool:
        with BackendManager(self.app, 'uia'):
            status_win = self.utils.get_window(app=self.app, title='???????????????????? ??????????????.+', regex=True)
            colvir_day = status_win['Static3'].window_text().strip()
        print(datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'), self.tomorrow, colvir_day, sep='    ')
        return self.tomorrow == colvir_day

    def run(self):
        try:
            pywinauto.Application(backend='win32').start(cmd_line=r'C:\CBS_R\COLVIR.exe')
            self.login()
            sleep(4)
        except (ElementNotFoundError, TimingsTimeoutError):
            self.retry()
            return
        try:
            self.pid: int = self.utils.get_current_process_pid(proc_name='COLVIR')
            self.app = pywinauto.Application(backend='win32').connect(process=self.pid)
            try:
                if self.app.Dialog.window_text() == '?????????????????? ????????????':
                    self.retry()
                    return
            except MatchError:
                pass
        except ProcessNotFoundError:
            sleep(1)
            self.pid: int = self.utils.get_current_process_pid(proc_name='COLVIR')
            self.app = pywinauto.Application(backend='win32').connect(process=self.pid)
        try:
            self.confirm_warning()
            sleep(1)
        except (ElementNotFoundError, MatchError):
            self.retry()
            return
        self.is_next_day = self.check_is_next_day()
        self.utils.kill_process(pid=self.pid)

    def retry(self) -> None:
        self.utils.kill_process(pid=self.pid)
        self.run()


def main():
    warnings.simplefilter(action='ignore', category=UserWarning)
    dotenv.load_dotenv()

    colvir_bot = Colvir(
        credentials=Credentials(usr=os.getenv('COLVIR_USR'), psw=os.getenv('COLVIR_PSW')),
        tomorrow='01.02.23',
    )

    with requests.Session() as session:
        notifier = TelegramNotifier(session=session)

        while True:
            colvir_bot.run()
            if colvir_bot.is_next_day:
                notifier.send_notification(message=f'!!!!!!!!')
                sleep(60)
                notifier.send_notification(message=f'!!!!!!!!')
                sleep(60)
                notifier.send_notification(message=f'!!!!!!!!')
                sleep(60)
                notifier.send_notification(message=f'!!!!!!!!')
                sleep(60)
                notifier.send_notification(message=f'!!!!!!!!')
                break
            else:
                sleep(60)


if __name__ == '__main__':
    main()
