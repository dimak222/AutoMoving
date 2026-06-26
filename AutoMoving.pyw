#-------------------------------------------------------------------------------
# Author:      dimak222
#
# Created:     25.06.2026
# Copyright:   (c) dimak222 2026
# Licence:     No
#-------------------------------------------------------------------------------

title = "AutoMoving"
ver = "v26.06.1"

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

# ---------- Настройки через QSettings ----------
settings = QSettings("AutoMoving", "AutoMoving")

if not settings.contains("move_enabled"):
    settings.setValue("move_enabled", False)        # по умолчанию перемещение выключено

if not settings.contains("interval"): # интервал обновления
    settings.setValue("interval", 300)

if not settings.contains("visual_move"):
    settings.setValue("visual_move", False)        # по умолчанию визуальный сдвиг выключен (с возвратом)

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
# Событие для управления потоком движения: установлено – перемещение активно, сброшено – поток спит
move_event = threading.Event()
# Инициализируем состояние события в соответствии с сохранённой настройкой
if settings.value("move_enabled", type=bool):
    move_event.set()
else:
    move_event.clear()

# Для обновления
pending_update = False
new_executable = None

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
def set_sleep_state(prevent: bool):
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

def manage_sleep_state():
    """
    Управляет состоянием предотвращения сна в зависимости от глобальной настройки
    и текущего состояния блокировки рабочего стола.
    Если опция включена, но компьютер заблокирован – сон разрешается.
    При разблокировке – снова запрещается.
    """
    locked = is_workstation_locked()
    prevent = settings.value("prevent_sleep", type=bool)
    set_sleep_state(prevent and not locked)

def toggle_prevent_sleep():
    """Переключает настройку предотвращения сна. Фактическое применение происходит в потоке движения."""
    current = settings.value("prevent_sleep", type=bool)
    settings.setValue("prevent_sleep", not current)
    # Немедленное применение не требуется, поток движения подхватит изменение за секунду.
    # Но для мгновенной реакции можно вызвать manage_sleep_state() здесь,
    # однако это может привести к двойному вызову из разных потоков, что не страшно.
    manage_sleep_state()

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

# ---------- Функции переключения ----------
def toggle_move_enabled():
    """Включает/выключает перемещение курсора и управляет событием move_event."""
    global move_event
    current = settings.value("move_enabled", type=bool)
    new_state = not current
    settings.setValue("move_enabled", new_state)
    if new_state:
        move_event.set()    # разрешаем движение
    else:
        move_event.clear()  # запрещаем движение (поток будет ждать)

def toggle_auto_update():
    current = settings.value("auto_update_check", type=bool)
    settings.setValue("auto_update_check", not current)

def toggle_visual_move():
    current = settings.value("visual_move", type=bool)
    settings.setValue("visual_move", not current)

# ---------- Движение мыши ----------
def move_mouse_randomly():

    # Для максимальной скорости движения мыши отключаем внутреннюю паузу pyautogui и fail-safe
    pyautogui.PAUSE = 0
    pyautogui.FAILSAFE = False

    global running, move_event
    directions = [(-5,-5), (-5,0), (-5,5), (0,-5), (0,5), (5,-5), (5,0), (5,5)]
    while running:
        # Регулярно актуализируем состояние предотвращения сна (раз в секунду)
        manage_sleep_state()

        # Ждём, пока перемещение не будет включено
        # move_event.wait() блокирует поток, но мы периодически просыпаемся,
        # чтобы проверить флаг running и состояние сна (раз в секунду)
        while not move_event.wait(timeout=1):
            if not running:
                return
            manage_sleep_state()   # проверка при каждом пробуждении

        # Если мы здесь, перемещение разрешено (move_event установлен)
        current_interval = settings.value("interval", type=int)
        # Сбрасываем счетчик и запоминаем текущую позицию перед началом отсчёта интервала
        slept = 0
        last_pos = pyautogui.position()

        # Ожидание с проверкой блокировки, смены интервала и активности пользователя
        while slept < current_interval:
            if not running:
                return
            # Если перемещение внезапно отключили – выходим во внешний цикл, где будем ждать событие
            if not settings.value("move_enabled", type=bool):
                break
            # Если интервал изменился во время ожидания – пересчитываем
            if settings.value("interval", type=int) != current_interval:
                break

            # При блокировке ПК только спим (без движения), но быстрее реагируем
            if is_workstation_locked():
                time.sleep(1)
                manage_sleep_state()  # сразу же обновляем сон, т.к. блокировка могла измениться
                slept += 1
                continue

            # Обычный отсчёт: спим по 1 секунде
            time.sleep(1)
            manage_sleep_state()   # проверяем блокировку после сна

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

        # Проверяем, что перемещение всё ещё включено и интервал не изменился
        if not settings.value("move_enabled", type=bool) or settings.value("interval", type=int) != current_interval:
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
    move_event.set()                # на всякий случай, чтобы поток не завис в wait()
    icon_obj.stop()                  # останавливаем pystray
    dialog_manager.quit_app.emit()  # сигналим Qt о завершении
    # При выходе сбрасываем состояние сна (возвращаем стандартное поведение)
    set_sleep_state(False)

def create_tray_icon():
    # Используем resource_path для загрузки иконки
    image = Image.open(resource_path('icon.ico'))

    # Вспомогательные лямбды для enabled/disabled
    def move_on():
        return settings.value("move_enabled", type=bool)

    menu = pystray.Menu(
        # Блок 1: Перемещение курсора
        pystray.MenuItem('Перемещение курсора',
                         toggle_move_enabled,
                         checked=lambda item: move_on()),

        pystray.MenuItem('Интервал', pystray.Menu(
            pystray.MenuItem('3 мин', lambda: set_interval(180),
                             checked=lambda item: settings.value("interval", type=int) == 180,
                             enabled=lambda item: move_on()),
            pystray.MenuItem('5 мин', lambda: set_interval(300),
                             checked=lambda item: settings.value("interval", type=int) == 300,
                             enabled=lambda item: move_on()),
            pystray.MenuItem('10 мин', lambda: set_interval(600),
                             checked=lambda item: settings.value("interval", type=int) == 600,
                             enabled=lambda item: move_on()),
            pystray.MenuItem('15 мин', lambda: set_interval(900),
                             checked=lambda item: settings.value("interval", type=int) == 900,
                             enabled=lambda item: move_on()),
            pystray.MenuItem('Свой...', custom_interval,
                             checked=lambda item: settings.value("interval", type=int) not in (180, 300, 600, 900),
                             enabled=lambda item: move_on()),
        ), enabled=lambda item: move_on()),

        pystray.MenuItem('Визуальный сдвиг',
                         toggle_visual_move,
                         checked=lambda item: settings.value("visual_move", type=bool),
                         enabled=lambda item: move_on()),

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

# ---------- Запуск ----------
if __name__ == "__main__":

    app = QApplication(sys.argv)
    app.setApplicationName(f"{title} {ver}")          # название приложения в системе
    app.setQuitOnLastWindowClosed(False)

    # Иконка приложения через ресурс
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
        # Автозакрытие через 4 секунды
        QTimer.singleShot(4000, msg.close)
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

    mover_thread = threading.Thread(target=move_mouse_randomly, daemon=True)
    mover_thread.start()

    icon = create_tray_icon()
    # Запускаем трей в отдельном потоке
    tray_thread = threading.Thread(target=icon.run, daemon=False)
    tray_thread.start()

    # Синхронизируем ярлык автозагрузки согласно настройке (с учётом .exe)
    sync_startup_shortcut()

    # Удаляем старый .old файл, оставшийся от предыдущего обновления
    cleanup_old_backup()

    # Запускаем главный цикл Qt (блокирует, пока не придёт сигнал quit)
    app.exec()

    # Запускаем новую версию, если было обновление
    if pending_update and new_executable:
        # Используем os.startfile для независимого запуска (не оставляет связей с текущим процессом)
        print(new_executable)
        os.startfile(new_executable)

    # Освобождаем мьютекс
    win32api.CloseHandle(mutex)