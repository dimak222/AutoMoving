#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     25.06.2026
# Copyright:   (c) dimak222 2026
# Licence:     No
#-------------------------------------------------------------------------------

title = "AutoMoving"
ver = "v26.07.1"

#------------------------------Импорт модулей-----------------------------------

import threading
import random
import time
import sys
import os
import ctypes
from pynput.mouse import Controller as MouseController
from pynput.keyboard import Key, Controller as KeyboardController
from PIL import Image
import pystray
from PyQt6.QtWidgets import QApplication, QInputDialog, QMessageBox
from PyQt6.QtCore import QObject, pyqtSignal, Qt, QSettings, QTimer
from PyQt6.QtGui import QIcon

import win32com.client
import win32event
import win32api
import winerror

import send2trash

from updater import start_update_check      # импортируем модуль обновления

# ---------- Функция для доступа к ресурсам (иконка внутри exe) ----------
def resource_path(relative_path):
    """Возвращает абсолютный путь к файлу, работая как в режиме разработки, так и в PyInstaller exe."""
    try:
        # PyInstaller сохраняет временную папку в _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(os.path.dirname(sys.argv[0]))
    return os.path.join(base_path, relative_path)

# ---------- Очистка старого .old файла при запуске ----------
def cleanup_old_backup():
    """Удаляет предыдущий .old файл, оставшийся после обновления."""
    old_path = sys.argv[0] + ".old"

    if not os.path.exists(old_path):
        return
    try:
        if send2trash:
            send2trash.send2trash(old_path)
        else:
            os.remove(old_path)
    except Exception:
        pass

# ---------- Очистка временного _new.exe после неудачного обновления ----------
def cleanup_new_exe():
    """Удаляет временный _new.exe файл, оставшийся после неудачного обновления."""
    if not is_exe():
        return
    current_exe = sys.argv[0]
    dir_name = os.path.dirname(current_exe)
    new_exe_path = os.path.join(dir_name, f"{title}_new.exe")
    if os.path.exists(new_exe_path):
        try:
            if send2trash:
                send2trash.send2trash(new_exe_path)
            else:
                os.remove(new_exe_path)
        except Exception:
            pass

# ---------- Настройки через QSettings ----------
settings = QSettings("AutoMoving", "AutoMoving")

if not settings.contains("move_enabled"):
    settings.setValue("move_enabled", False)        # по умолчанию перемещение выключено

if not settings.contains("interval"): # интервал обновления
    settings.setValue("interval", 300)

if not settings.contains("visual_move"):
    settings.setValue("visual_move", False)        # по умолчанию визуальный сдвиг выключен (с возвратом)

if not settings.contains("keyboard_simulation"):
    settings.setValue("keyboard_simulation", False) # по умолчанию имитация клавишей выключена

if not settings.contains("prevent_sleep"):
    settings.setValue("prevent_sleep", True)       # по умолчанию предотвращаем сон

if not settings.contains("auto_update_check"):
    settings.setValue("auto_update_check", True)   # по умолчанию автообновление включено

if not settings.contains("auto_startup"):
    settings.setValue("auto_startup", True)        # по умолчанию автозагрузка включена

# ---------- Глобальные переменные ----------
running = True
dialog_manager = None
icon = None

# Контроллеры pynput
mouse = MouseController()
keyboard = KeyboardController()

# Для обновления
pending_update = False
new_executable = None

# ---------- Определение времени последнего ввода пользователя ----------
class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint),
                ("dwTime", ctypes.c_uint)]

def get_idle_time():
    """Возвращает время бездействия пользователя в секундах (мышь и клавиатура)."""
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
        millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
        return millis / 1000.0
    return 0

# ---------- Проверка блокировки рабочего стола ----------
def is_workstation_locked():
    """Возвращает True, если рабочий стол заблокирован (Win+L)."""
    # Пытаемся открыть активный рабочий стол ввода
    hdesk = ctypes.windll.user32.OpenInputDesktop(0, False, 0)
    if hdesk == 0:
        # Не удалось открыть – скорее всего, сессия заблокирована
        return True
    # Успешно открыли – сразу закрываем, чтобы не утекал дескриптор
    ctypes.windll.user32.CloseDesktop(hdesk)
    return False

# ---------- Предотвращение сна / выключения экрана ----------
def set_sleep_state(prevent: bool): # запрещаем/разрешаем системе переходить в спящий режим
    """
    Если prevent == True, запрещаем системе переходить в спящий режим
    и выключать дисплей. Иначе сбрасываем флаг.
    """
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    if prevent:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )
    else:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

# ---------- Работа с автозагрузкой ----------
def get_startup_shortcut_path():
    """Возвращает путь к ярлыку в папке автозагрузки текущего пользователя."""
    startup_folder = os.path.join(os.getenv('APPDATA'),
                                  'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    return os.path.join(startup_folder, f'{title}.lnk')

def is_exe():
    """Проверяет, является ли текущий исполняемый файл .exe."""
    target = os.path.abspath(sys.argv[0])
    return os.path.splitext(target)[1].lower() == '.exe'

def sync_startup_shortcut():
    """Синхронизирует наличие ярлыка автозагрузки с настройкой auto_startup."""
    shortcut_path = get_startup_shortcut_path()
    enabled = settings.value("auto_startup", type=bool)
    if enabled and is_exe() and not os.path.exists(shortcut_path):
        # Создать ярлык
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            target = os.path.abspath(sys.argv[0])
            working_dir = os.path.dirname(target)
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target
            shortcut.WorkingDirectory = working_dir
            shortcut.Description = f"{title} {ver}"
            shortcut.IconLocation = target
            shortcut.Save()
        except ImportError:
            pass
    elif (not enabled) and os.path.exists(shortcut_path):
        # Удалить ярлык
        try:
            os.remove(shortcut_path)
        except OSError:
            pass

# ---------- Qt-прослойка с сигналом завершения ----------
class IntervalDialog(QObject):
    request_dialog = pyqtSignal()
    quit_app = pyqtSignal()          # <-- сигнал для завершения приложения

    def __init__(self):
        super().__init__()
        self.request_dialog.connect(self.show_dialog, Qt.ConnectionType.QueuedConnection)
        self.seconds = None
        self.done = False

    def trigger(self):
        self.done = False
        self.seconds = None
        self.request_dialog.emit()

    def show_dialog(self):
        num, ok = QInputDialog.getInt(
            None, f"{title} {ver}", "Введите интервал движения в секундах:",  # ← заголовок с версией
            value=settings.value("interval", type=int), min=1
        )
        if ok:
            self.seconds = num
        else:
            self.seconds = None
        self.done = True

# ---------- Универсальная функция переключения булевых настроек ----------
def toggle_setting(key):
    """Инвертирует булеву настройку."""
    current = settings.value(key, type=bool)
    settings.setValue(key, not current)
    return not current          # возвращаем новое значение

# ---------- Отдельные функции для меню ----------
def toggle_move_enabled():
    toggle_setting("move_enabled")

def toggle_auto_update():
    toggle_setting("auto_update_check")

def toggle_visual_move():
    toggle_setting("visual_move")

def toggle_keyboard_simulation():
    toggle_setting("keyboard_simulation")

def toggle_prevent_sleep():
    toggle_setting("prevent_sleep")

def toggle_startup():
    """Переключает автозагрузку и синхронизирует ярлык."""
    current = settings.value("auto_startup", type=bool)
    settings.setValue("auto_startup", not current)
    sync_startup_shortcut()

# ---------- Установка интервала  ----------
def set_interval(seconds):
    settings.setValue("interval", seconds)

def custom_interval():
    global dialog_manager
    if dialog_manager is None:
        return
    dialog_manager.trigger()
    for _ in range(600):
        if dialog_manager.done:
            break
        time.sleep(0.05)
        QApplication.instance().processEvents()
    if dialog_manager.done and dialog_manager.seconds is not None:
        set_interval(dialog_manager.seconds)

# ---------- Перемещение курсора ----------
def cursor_movement():
    """Передвижение курсора."""
    directions = [(-5,-5), (-5,0), (-5,5), (0,-5), (0,5), (5,-5), (5,0), (5,5)]
    dx, dy = random.choice(directions)
    mouse.move(dx, dy)
    # Визуальный сдвиг: если выключен, возвращаем курсор; иначе оставляем на новом месте
    if not settings.value("visual_move", type=bool):
        mouse.move(-dx, -dy) # мгновенный возврат на исходную позицию

# ---------- Эмуляция нажатия клавиши Ctrl через pynput ----------
def simulate_key_press():
    """Отправляет нажатие и отпускание клавиши Ctrl."""
    keyboard.press(Key.ctrl)
    keyboard.release(Key.ctrl)

# ---------- Основной рабочий поток ----------
def run():
    global running

    last_action_time = 0         # Внутренний таймер последнего действия программы
    was_locked = False           # Триггер первого входа в блокировку

    while running:

        locked = is_workstation_locked() # Проверка блокировки и управление сном

        if locked:
            if not was_locked:
                set_sleep_state(False) # при входе в блокировку – отключаем предотвращение сна

        was_locked = locked

        current_interval = settings.value("interval", type=int)
        now = time.time()

        # Проверяем, пора ли выполнять действия (внутренний интервал + бездействие пользователя)
        if now - last_action_time >= current_interval and get_idle_time() >= current_interval:
            if not locked:
                if settings.value("move_enabled", type=bool):
                    cursor_movement()
                if settings.value("keyboard_simulation", type=bool):
                    simulate_key_press()
                if settings.value("prevent_sleep", type=bool):
                    set_sleep_state(True)

                last_action_time = now

        time.sleep(1)   # небольшая пауза

# ---------- Трей ----------
def stop_program(icon_obj):
    global running, dialog_manager
    running = False
    icon_obj.stop()               # останавливаем pystray
    dialog_manager.quit_app.emit()# сигналим Qt о завершении
    set_sleep_state(False)        # сбрасываем предотвращение сна

def create_tray_icon():
    # Используем resource_path для загрузки иконки
    image = Image.open(resource_path('icon.ico'))

    # Вспомогательная лямбда для пунктов, зависящих от move_enabled
    def move_on():
        return settings.value("move_enabled", type=bool)

    menu = pystray.Menu(
        # Блок 1: Перемещение курсора
        pystray.MenuItem('Перемещение курсора',
                         toggle_move_enabled,
                         checked=lambda item: move_on()),

        pystray.MenuItem('Визуальный сдвиг',
                         toggle_visual_move,
                         checked=lambda item: settings.value("visual_move", type=bool),
                         enabled=lambda item: move_on()),

        pystray.MenuItem('Имитация клавиши (Ctrl)',
                         toggle_keyboard_simulation,
                         checked=lambda item: settings.value("keyboard_simulation", type=bool)),

        pystray.MenuItem('Интервал', pystray.Menu(
            pystray.MenuItem('3 мин', lambda: set_interval(180),
                             checked=lambda item: settings.value("interval", type=int) == 180),
            pystray.MenuItem('5 мин', lambda: set_interval(300),
                             checked=lambda item: settings.value("interval", type=int) == 300),
            pystray.MenuItem('10 мин', lambda: set_interval(600),
                             checked=lambda item: settings.value("interval", type=int) == 600),
            pystray.MenuItem('15 мин', lambda: set_interval(900),
                             checked=lambda item: settings.value("interval", type=int) == 900),
            pystray.MenuItem('Свой...', custom_interval,
                             checked=lambda item: settings.value("interval", type=int) not in (180, 300, 600, 900)),
        )),

        pystray.Menu.SEPARATOR,

        # Блок 2: Предотвращать сон
        pystray.MenuItem('Предотвращать сон',
                         toggle_prevent_sleep,
                         checked=lambda item: settings.value("prevent_sleep", type=bool)),

        pystray.Menu.SEPARATOR,

        # Блок 3: Обновления и автозагрузка
        pystray.MenuItem('Проверять обновления при старте',
                         toggle_auto_update,
                         checked=lambda item: settings.value("auto_update_check", type=bool)),

        pystray.MenuItem('Запуск при старте системы',
                         toggle_startup,
                         checked=lambda item: settings.value("auto_startup", type=bool)),

        pystray.Menu.SEPARATOR,

        pystray.MenuItem('Выйти из программы', lambda: stop_program(icon))
    )
    # Название в трее (всплывающая подсказка) и идентификатор с версией
    return pystray.Icon(f"{title} {ver}", image, f"{title} {ver}", menu)

# ---------- Обработчик завершения обновления ----------
def restart_program(new_exe):
    """Вызывается из updater, когда всё готово к перезапуску."""
    global pending_update, new_executable
    pending_update = True
    new_executable = new_exe
    stop_program(icon)

def create_update_bat(new_exe_path):
    """Создаёт bat-файл в TEMP, который запустит новую версию после выхода из текущей."""
    import tempfile
    bat_path = os.path.join(tempfile.gettempdir(), f"{title}_update.bat")
    with open(bat_path, "w") as f:
        f.write(f"""@echo off
timeout /t 2 /nobreak >nul
start "" "{new_exe_path}"
del "%~f0" & exit
""")
    return bat_path

# ---------- Запуск ----------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(f"{title} {ver}")
    app.setQuitOnLastWindowClosed(False)

    # Иконка приложения через ресурс (для exe файла)
    icon_path = resource_path('icon.ico')

    # Проверка на повторный запуск (используем локальный мьютекс)
    mutex_name = f"Local\\{title}SingleInstance"
    mutex = win32event.CreateMutex(None, False, mutex_name)
    last_error = win32api.GetLastError()
    if last_error == winerror.ERROR_ALREADY_EXISTS:
        msg = QMessageBox()
        msg.setWindowTitle(f"{title} {ver}")
        msg.setText("Приложение уже запущено!")
        msg.setIcon(QMessageBox.Icon.Warning)
        if os.path.exists(icon_path):
            msg.setWindowIcon(QIcon(icon_path))
        QTimer.singleShot(4000, msg.close) # Автозакрытие через 4 секунды
        msg.setWindowFlags(msg.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        msg.exec()
        win32api.CloseHandle(mutex)
        sys.exit(0)

    # Автоматическая проверка обновлений при старте, только если включена в настройках
    if settings.value("auto_update_check", type=bool):
        _update_thread = start_update_check(
            title, ver, None,
            log_callback=lambda msg: print(msg),
            on_restart=restart_program
        )
    else:
        print("Автопроверка обновлений отключена.")

    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    dialog_manager = IntervalDialog()
    # Подключаем сигнал завершения к выходу из Qt
    dialog_manager.quit_app.connect(app.quit, Qt.ConnectionType.QueuedConnection)

    run_thread = threading.Thread(target=run, daemon=True)
    run_thread.start()

    icon = create_tray_icon()

    tray_thread = threading.Thread(target=icon.run, daemon=False) # Запускаем трей в отдельном потоке
    tray_thread.start()

    sync_startup_shortcut() # Синхронизируем ярлык автозагрузки согласно настройке (с учётом .exe)
    cleanup_old_backup() # Удаляем старый .old файл, оставшийся от предыдущего обновления
    cleanup_new_exe()       # Удаляем временный _new.exe, если остался после сбоя

    app.exec()

    win32api.CloseHandle(mutex) # Освобождаем мьютекс

    if pending_update and new_executable: # Запускаем новую версию, если было обновление
        bat = create_update_bat(new_executable)
        os.startfile(bat) # Используем os.startfile для независимого запуска (не оставляет связей с текущим процессом)