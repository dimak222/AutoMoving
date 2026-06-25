#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     25.06.2026
# Copyright:   (c) dimak222 2026
# Licence:     No
#-------------------------------------------------------------------------------

title = "AutoMoving"
ver = "v26.06.0"

#------------------------------Импорт модулей-----------------------------------

import threading
import random
import time
import sys
import os
import ctypes
import pyautogui
from PIL import Image
import pystray
from PyQt6.QtWidgets import QApplication, QInputDialog, QMessageBox
from PyQt6.QtCore import QObject, pyqtSignal, Qt, QSettings, QTimer
from PyQt6.QtGui import QIcon

import win32com.client
import win32event
import win32api
import winerror

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

# ---------- Настройки через QSettings ----------
settings = QSettings("AutoMoving", "AutoMoving")
if not settings.contains("interval"):
    settings.setValue("interval", 600)

if not settings.contains("visual_move"):
    settings.setValue("visual_move", False)        # по умолчанию визуальный сдвиг выключен (с возвратом)

if not settings.contains("auto_update_check"):
    settings.setValue("auto_update_check", True)   # по умолчанию автообновление включено

if not settings.contains("auto_startup"):
    settings.setValue("auto_startup", True)        # по умолчанию автозагрузка включена

# ---------- Глобальные переменные ----------
running = True
dialog_manager = None
icon = None

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

# ---------- Работа с автозагрузкой (только Windows) ----------
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
    # Если не .exe и настройка включена — ярлык не создаётся (пользователь увидит,
    # что галочка есть, но ярлыка нет; можно добавить уведомление, но оставим так)

def toggle_startup():
    """Переключает настройку автозагрузки и синхронизирует ярлык."""
    current = settings.value("auto_startup", type=bool)
    settings.setValue("auto_startup", not current)
    sync_startup_shortcut()

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

# ---------- Установка интервала ----------
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

# ---------- Функция переключения автообновления ----------
def toggle_auto_update():
    current = settings.value("auto_update_check", type=bool)
    settings.setValue("auto_update_check", not current)

# ---------- Функция переключения визуального сдвига ----------
def toggle_visual_move():
    current = settings.value("visual_move", type=bool)
    settings.setValue("visual_move", not current)

# ---------- Движение мыши ----------
def move_mouse_randomly():

    # Для максимальной скорости движения мыши отключаем внутреннюю паузу pyautogui и fail-safe
    pyautogui.PAUSE = 0
    pyautogui.FAILSAFE = False

    global running
    directions = [(-1,-1), (-1,0), (-1,1), (0,-1), (0,1), (1,-1), (1,0), (1,1)]
    while running:
        current_interval = settings.value("interval", type=int)
        # Сбрасываем счетчик и запоминаем текущую позицию перед началом отсчёта интервала
        slept = 0
        last_pos = pyautogui.position()

        # Ожидание с проверкой блокировки, смены интервала и активности пользователя
        while slept < current_interval:
            if not running:
                return
            # Если интервал изменился во время ожидания – пересчитываем
            if settings.value("interval", type=int) != current_interval:
                break
            # При блокировке ПК только спим (без движения), но быстрее реагируем
            if is_workstation_locked():
                time.sleep(1)
                slept += 1
                continue

            # Обычный отсчёт: спим по 1 секунде
            time.sleep(1)

            # Проверяем, двигал ли пользователь мышь (или что-то ещё) во время ожидания
            cur_pos = pyautogui.position()
            if cur_pos != last_pos:
                # Было перемещение – сбрасываем таймер и запоминаем новую позицию
                slept = 0
                last_pos = cur_pos
            else:
                slept += 1

        if not running:
            return

        # Проверяем, не изменился ли интервал во время сна
        if settings.value("interval", type=int) != current_interval:
            continue

        # Если система не заблокирована, выполняем движение
        if not is_workstation_locked():
            dx, dy = random.choice(directions)

            pyautogui.move(dx, dy)      # сдвиг на 1 пиксель

            # Визуальный сдвиг: если выключен, возвращаем курсор; иначе оставляем на новом месте
            if not settings.value("visual_move", type=bool):
                pyautogui.move(-dx, -dy)    # мгновенный возврат на исходную позицию

# ---------- Трей ----------
def stop_program(icon_obj):
    global running, dialog_manager
    running = False
    icon_obj.stop()                  # останавливаем pystray
    dialog_manager.quit_app.emit()  # сигналим Qt о завершении

def create_tray_icon():
    # Используем resource_path для загрузки иконки
    image = Image.open(resource_path('icon.ico'))
    menu = pystray.Menu(
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

        pystray.MenuItem('Визуальный сдвиг',
                         toggle_visual_move,
                         checked=lambda item: settings.value("visual_move", type=bool)),

        pystray.MenuItem('Проверять обновления при старте',
                         toggle_auto_update,
                         checked=lambda item: settings.value("auto_update_check", type=bool)),

        pystray.MenuItem('Запуск при старте системы',
                         toggle_startup,
                         checked=lambda item: settings.value("auto_startup", type=bool)),

        pystray.MenuItem('Выйти из программы', lambda: stop_program(icon))
    )
    # Название в трее (всплывающая подсказка) и идентификатор с версией
    return pystray.Icon(f"{title} {ver}", image, f"{title} {ver}", menu)

# ---------- Запуск ----------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(f"{title} {ver}")          # название приложения в системе
    app.setQuitOnLastWindowClosed(False)

    # Иконка приложения через ресурс
    icon_path = resource_path('icon.ico')

    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

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
        # Автозакрытие через 4 секунды
        QTimer.singleShot(4000, msg.close)
        msg.exec()
        win32api.CloseHandle(mutex)
        sys.exit(0)

    dialog_manager = IntervalDialog()
    # Подключаем сигнал завершения к выходу из Qt
    dialog_manager.quit_app.connect(app.quit, Qt.ConnectionType.QueuedConnection)

    mover_thread = threading.Thread(target=move_mouse_randomly, daemon=True)
    mover_thread.start()

    icon = create_tray_icon()
    # Запускаем трей в отдельном потоке
    tray_thread = threading.Thread(target=icon.run, daemon=False)
    tray_thread.start()

    # Синхронизируем ярлык автозагрузки согласно настройке (с учётом .exe)
    sync_startup_shortcut()

    # Автоматическая проверка обновлений при старте, только если включена в настройках
    if settings.value("auto_update_check", type=bool):
        _update_thread = start_update_check(title, ver, None,
                                            log_callback=lambda msg: print(msg))
    else:
        print("Автопроверка обновлений отключена.")

    # Запускаем главный цикл Qt (блокирует, пока не придёт сигнал quit)
    app.exec()

    # Дополнительная гарантия: ждём завершения потока трея (не обязательно, но для порядка)
    tray_thread.join(timeout=2)

    # Освобождаем мьютекс
    win32api.CloseHandle(mutex)