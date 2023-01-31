import os
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


class Colvir:
    def __init__(self, credentials: Credentials, today: str, session: requests.Session) -> None:
        self.credentials = credentials
        self.pid: int or None = None
        self.app: pywinauto.Application or None = None
        self.today: str = today
        self.notifier = TelegramNotifier(session=session)
        self.utils = Utils()

    def login(self):
        desktop = pywinauto.Desktop(backend='win32')
        try:
            login_win = desktop.window(title='Вход в систему')
            login_win.wait(wait_for='exists', timeout=20)
            login_win['Edit2'].wrapper_object().set_text(text=self.credentials.usr)
            login_win['Edit'].wrapper_object().set_text(text=self.credentials.psw)
            login_win['OK'].wrapper_object().click()
        except ElementAmbiguousError:
            windows: List[DialogWrapper] = pywinauto.Desktop(backend='win32').windows()
            for win in windows:
                if 'Вход в систему' not in win.window_text():
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

    def is_next_day(self) -> bool:
        with BackendManager(self.app, 'uia'):
            status_win = self.utils.get_window(app=self.app, title='Банковская система.+', regex=True)
            colvir_day = status_win['Static3'].window_text().strip()
        return self.today != colvir_day

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
                if self.app.Dialog.window_text() == 'Произошла ошибка':
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
        if self.is_next_day():
            self.notifier.send_notification(message=f'!!!!!!!!')

    def retry(self) -> None:
        self.utils.kill_process(pid=self.pid)
        self.run()


def main():
    dotenv.load_dotenv()

    with requests.Session() as session:
        colvir_bot = Colvir(
            credentials=Credentials(usr=os.getenv('COLVIR_USR'), psw=os.getenv('COLVIR_PSW')),
            today='31.01.22',
            session=session
        )
        colvir_bot.run()


if __name__ == '__main__':
    main()
