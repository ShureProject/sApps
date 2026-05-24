import sys
import os
import subprocess
import requests
import json
import shutil
import winshell
from win32com.client import Dispatch
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QCheckBox, QPushButton, QProgressBar,
                             QLabel, QScrollArea, QFrame, QMessageBox, QGridLayout,
                             QDialog, QDialogButtonBox, QStyle, QStyleOptionButton,
                             QSizePolicy, QSpacerItem, QTabWidget)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl, QSize, QTimer, QRect, QProcess
from PyQt6.QtGui import QFont, QPalette, QColor, QIcon, QPixmap, QFontDatabase, QPainter, QPen, QBrush
import tempfile
import ctypes

class DownloadThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, str)
    error = pyqtSignal(str)
    current_file = pyqtSignal(str)

    def __init__(self, url, save_path, app_name):
        super().__init__()
        self.url = url
        self.save_path = save_path
        self.app_name = app_name

    def run(self):
        try:
            self.current_file.emit(self.app_name)
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.5,en;q=0.3',
                'Connection': 'keep-alive',
            }
            response = requests.get(self.url, stream=True, headers=headers, allow_redirects=True, timeout=30)
            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))
            block_size = 8192
            downloaded = 0
            with open(self.save_path, 'wb') as file:
                for data in response.iter_content(block_size):
                    downloaded += len(data)
                    file.write(data)
                    if total_size > 0:
                        progress_percent = int((downloaded / total_size) * 100)
                        self.progress.emit(progress_percent)
                    else:
                        self.progress.emit(-1)
            self.finished.emit(self.save_path, self.app_name)
        except requests.exceptions.RequestException as e:
            self.error.emit(f"Ошибка сети: {str(e)}")
        except Exception as e:
            self.error.emit(f"Ошибка: {str(e)}")


class InitWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("sApps - Загрузка")
        self.setFixedSize(400, 150)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        container = QWidget()
        container.setStyleSheet("""
            QWidget {
                background-color: #2a2a2a;
                border: 2px solid #4a9eff;
                border-radius: 12px;
            }
        """)
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(20, 20, 20, 20)
        container_layout.setSpacing(15)

        title = QLabel("sApps")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #4a9eff; border: none; background: transparent;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        container_layout.addWidget(title)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(6)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 3px;
                background-color: #3a3a3a;
            }
            QProgressBar::chunk {
                background-color: #4a9eff;
                border-radius: 3px;
            }
        """)
        container_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Подготовка...")
        self.status_label.setFont(QFont("Segoe UI", 9))
        self.status_label.setStyleSheet("color: #7a7a7a; border: none; background: transparent;")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        container_layout.addWidget(self.status_label)

        layout.addWidget(container)
        self.center()

    def center(self):
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

    def set_status(self, text):
        self.status_label.setText(text)
        QApplication.processEvents()


class IconLoader(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, icon_dir, programs):
        super().__init__()
        self.icon_dir = icon_dir
        self.programs = programs

    def run(self):
        session = requests.Session()
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
            'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
            'Referer': 'https://www.google.com/',
        })

        for category_name, category_programs in self.programs.items():
            for program in category_programs:
                app_name = program["name"]
                cache_path = os.path.join(self.icon_dir, f"{app_name.replace(' ', '_')}.png")

                if os.path.exists(cache_path):
                    continue

                self.progress.emit(f"Загрузка иконки: {app_name}")

                icon_url = program["icon_url"]
                if "upload.wikimedia.org" in icon_url and "/thumb/" in icon_url:
                    parts = icon_url.split("/")
                    if "thumb" in parts:
                        thumb_index = parts.index("thumb")
                        del parts[thumb_index]
                        del parts[-1]
                    icon_url = "/".join(parts)

                try:
                    response = session.get(icon_url, timeout=10, allow_redirects=True)
                    if response.status_code == 200:
                        pixmap = QPixmap()
                        pixmap.loadFromData(response.content)
                        if not pixmap.isNull():
                            pixmap.save(cache_path, "PNG")
                except:
                    pass

                QThread.msleep(100)

        self.finished.emit()

class SuccessDialog(QDialog):
    def __init__(self, app_name, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.parent_window = parent
        self.setWindowTitle("Загрузка завершена")
        self.setModal(True)
        self.setFixedSize(450, 220)

        self.setStyleSheet("""
            QDialog {
                background-color: #2a2a2a;
                border: 2px solid #4a9eff;
                border-radius: 10px;
            }
            QLabel {
                color: #e0e0e0;
                font-size: 14px;
            }
            QPushButton {
                background-color: #4a9eff;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #3a8eef;
            }
            QPushButton#cancelButton {
                background-color: #5a5a5a;
            }
            QPushButton#cancelButton:hover {
                background-color: #6a6a6a;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        success_label = QLabel(f"✅ {app_name} успешно скачан!")
        success_label.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        success_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(success_label)

        info_label = QLabel(f"Файл сохранен в:\n{file_path}")
        info_label.setFont(QFont("Segoe UI", 10))
        info_label.setStyleSheet("color: #8a8a8a;")
        info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(15)

        self.launch_button = QPushButton("🚀 Запустить")
        self.launch_button.clicked.connect(self.launch_app)

        self.close_button = QPushButton("Закрыть")
        self.close_button.setObjectName("cancelButton")
        self.close_button.clicked.connect(self.accept)

        buttons_layout.addWidget(self.launch_button)
        buttons_layout.addWidget(self.close_button)

        layout.addLayout(buttons_layout)

    def launch_app(self):
        try:
            if os.path.exists(self.file_path):
                print(f"Запуск приложения: {self.file_path}")
                if sys.platform == 'win32':
                    subprocess.Popen([self.file_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                else:
                    subprocess.Popen([self.file_path])
                self.hide()
                QTimer.singleShot(500, self.accept)
        except Exception as e:
            print(f"Ошибка запуска: {str(e)}")
            QMessageBox.warning(self, "Ошибка", f"Не удалось запустить приложение:\n{str(e)}")
            self.accept()

    def closeEvent(self, event):
        event.accept()

class CustomCheckBox(QWidget):
    stateChanged = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(20, 20)
        self.checked = False
        self.enabled = True

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRect(0, 0, 19, 19)
        if self.checked:
            painter.setBrush(QBrush(QColor("#4a9eff")))
            painter.setPen(QPen(QColor("#4a9eff"), 2))
            painter.drawRoundedRect(rect, 3, 3)
            painter.setPen(QPen(QColor("white"), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            left_point = QRect(5, 8, 4, 8).center()
            bottom_point = QRect(8, 11, 4, 8).center()
            right_point = QRect(13, 5, 7, 7).center()
            painter.drawLine(left_point, bottom_point)
            painter.drawLine(bottom_point, right_point)
        else:
            if self.enabled:
                painter.setBrush(QBrush(QColor("#3a3a3a")))
                painter.setPen(QPen(QColor("#4a9eff"), 1.5))
            else:
                painter.setBrush(QBrush(QColor("#2a2a2a")))
                painter.setPen(QPen(QColor("#4a4a4a"), 1.5))
            painter.drawRoundedRect(rect, 3, 3)

    def mousePressEvent(self, event):
        if self.enabled:
            self.setChecked(not self.checked)
            self.stateChanged.emit(self.checked)

    def setChecked(self, checked):
        self.checked = checked
        self.update()

    def isChecked(self):
        return self.checked

    def setEnabled(self, enabled):
        self.enabled = enabled
        self.update()

class ProgramCard(QFrame):
    clicked = pyqtSignal()
    doubleClicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.checked = False
        self.is_downloaded = False
        self.setFrameStyle(QFrame.Shape.StyledPanel)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.setMinimumHeight(65)
        self.setMaximumHeight(72)
        self.setMinimumWidth(280)
        self.setMaximumWidth(350)

    def mousePressEvent(self, event):
        if not self.is_downloaded:
            self.clicked.emit()
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        if self.is_downloaded:
            self.doubleClicked.emit()
        super().mouseDoubleClickEvent(event)

class DownloaderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.download_threads = []
        self.downloaded_apps = {}
        self.apps_dir = os.path.join(os.path.expanduser("~"), "sApps")
        self.downloads_file = os.path.join(self.apps_dir, "downloads.txt")
        self.icon_dir = os.path.join(self.apps_dir, "icons")
        self.program_cards = []
        self.is_downloading = False

        self.create_apps_directory()
        self.load_downloaded_apps()

        self.setVisible(False)
        self.hide()

        self.init_window = InitWindow()
        self.init_window.show()
        QApplication.processEvents()

        self.load_icons_async()

    def load_icons_async(self):
        self.init_window.set_status("Загрузка иконок...")
        self.icon_loader = IconLoader(self.icon_dir, self.get_programs())
        self.icon_loader.progress.connect(self.init_window.set_status)
        self.icon_loader.finished.connect(self.on_icons_loaded)
        self.icon_loader.start()

    def on_icons_loaded(self):
        self.init_window.set_status("Запуск приложения...")

        self.init_ui()
        self.setup_programs()

        QTimer.singleShot(200, self.finish_init)

    def finish_init(self):
        self.setVisible(True)
        self.show()
        self.raise_()
        self.activateWindow()
        self.setWindowState(Qt.WindowState.WindowActive)

        QTimer.singleShot(100, self.close_init_window)

    def close_init_window(self):
        if hasattr(self, 'init_window') and self.init_window:
            self.init_window.close()
            self.init_window.deleteLater()

    def create_apps_directory(self):
        try:
            if not os.path.exists(self.apps_dir):
                os.makedirs(self.apps_dir)
                print(f"Создана папка: {self.apps_dir}")
            if not os.path.exists(self.icon_dir):
                os.makedirs(self.icon_dir)
                print(f"Создана папка для иконок: {self.icon_dir}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось создать папку {self.apps_dir}: {str(e)}")
            sys.exit(1)

    def create_start_menu_shortcut(self, app_name, target_path):
        try:
            startup_folder = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
            if not os.access(startup_folder, os.W_OK):
                startup_folder = os.path.join(os.path.expanduser("~"),
                                              "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs")
            shortcut_path = os.path.join(startup_folder, f"{app_name}.lnk")
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = target_path
            shortcut.WorkingDirectory = os.path.dirname(target_path)
            shortcut.IconLocation = target_path
            shortcut.save()
            print(f"Ярлык создан: {shortcut_path}")
            return True
        except Exception as e:
            print(f"Ошибка создания ярлыка для {app_name}: {str(e)}")
            return False

    def load_downloaded_apps(self):
        try:
            if os.path.exists(self.downloads_file):
                with open(self.downloads_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and " - " in line:
                            app_name, file_path = line.split(" - ", 1)
                            file_path = file_path.strip().strip('"')
                            if os.path.exists(file_path):
                                self.downloaded_apps[app_name] = file_path
        except Exception as e:
            print(f"Ошибка загрузки списка: {str(e)}")
            self.downloaded_apps = {}

    def save_downloaded_app(self, app_name, file_path):
        try:
            self.remove_app_from_file(app_name)
            with open(self.downloads_file, 'a', encoding='utf-8') as f:
                f.write(f'{app_name} - "{file_path}"\n')
            self.downloaded_apps[app_name] = file_path
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось сохранить информацию: {str(e)}")

    def remove_app_from_file(self, app_name):
        try:
            if os.path.exists(self.downloads_file):
                with open(self.downloads_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                with open(self.downloads_file, 'w', encoding='utf-8') as f:
                    for line in lines:
                        if not line.startswith(f"{app_name} - "):
                            f.write(line)
        except Exception as e:
            print(f"Ошибка удаления записи: {str(e)}")

    def is_app_downloaded(self, app_name):
        if app_name in self.downloaded_apps:
            file_path = self.downloaded_apps[app_name]
            if os.path.exists(file_path):
                return True
            else:
                del self.downloaded_apps[app_name]
                self.remove_app_from_file(app_name)
                return False
        return False

    def launch_app(self, app_name):
        if app_name in self.downloaded_apps:
            file_path = self.downloaded_apps[app_name]
            if os.path.exists(file_path):
                try:
                    if sys.platform == 'win32':
                        subprocess.Popen([file_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    else:
                        subprocess.Popen([file_path])
                    return True
                except Exception as e:
                    QMessageBox.warning(self, "Ошибка", f"Не удалось запустить {app_name}:\n{str(e)}")
            else:
                del self.downloaded_apps[app_name]
                self.remove_app_from_file(app_name)
                self.setup_programs()
        return False

    def get_programs(self):
        return {
            "Основные": [
                {"name": "Legacy Launcher", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRuAeuk1fSDVDS504lNdFQ3su_qzOqZ7pk5MQ&s", "download_url": "https://dl.legacylauncher.ru/legacy/installer"},
                {"name": "Chrome", "icon_url": "https://www.google.com/chrome/static/images/chrome-logo-m100.svg", "download_url": "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B1B1614C2-0383-07AE-1532-223816242268%7D%26lang%3Dru%26browser%3D4%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3D-arch_x64-statsdef_1%26installdataindex%3Dempty/update2/installers/ChromeSetup.exe"},
                {"name": "FireFox", "icon_url": "https://store-images.s-microsoft.com/image/apps.7279.14473293538384797.bcb417dc-ffbe-444e-9589-e6a25f04ad52.156eed19-aa35-4e69-96a7-c11abd7f887d", "download_url": "https://www.firefox.com/thanks/"},
                {"name": "Steam", "icon_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/8/83/Steam_icon_logo.svg/3840px-Steam_icon_logo.svg.png", "download_url": "https://cdn.fastly.steamstatic.com/client/installer/SteamSetup.exe"},
                {"name": "YouTube Music", "icon_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/6/6a/Youtube_Music_icon.svg/960px-Youtube_Music_icon.svg.png", "download_url": "https://www.dropbox.com/scl/fi/t20modbnx3uz3rfgxl0s9/YouTube-Music.exe?rlkey=dfua9z4h5ph6umzui3qsnejo1&st=31zpkycn&dl=1"},
                {"name": "Discord", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcR8aMugg7LWDXqkWc-9JlApM4MLPXhi-EPDYA&s", "download_url": "https://www.softportal.com/getsoft-44158-discord-100.html"},
                {"name": "Telegram", "icon_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/8/83/Telegram_2019_Logo.svg/250px-Telegram_2019_Logo.svg.png", "download_url": "https://dl.comss.org/download/tsetup.6.8.2.exe"},
                {"name": "WinRAR", "icon_url": "https://www.win-rar.com/fileadmin/images/winrar-archive.png", "download_url": "https://www.win-rar.com/fileadmin/winrar-versions/winrar/winrar-x64-722ru.exe"},
                {"name": "7-Zip", "icon_url": "https://www.7-zip.org/7ziplogo.png", "download_url": "https://github.com/ip7z/7zip/releases/download/26.01/7z2601-x64.exe"},
                {"name": "Notepad++", "icon_url": "https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEhZqrweT9UQpLoiyknOmv1OYEcBYH-RMTBaS5N9SS-_eVxsoxrnEuHvH0HmvPVtIBneMc9iXveD2CaPR-Rhf1q8cFlqjs95r5VEhYZ1vPHPgXWom2tZVv6uH7wSfk7CX-3ZOKm3XLqCIlA/s1600/notepad-plus-plus.png", "download_url": "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.9.6/npp.8.9.6.Installer.x64.exe"},
                {"name": "VS Code", "icon_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/9/9a/Visual_Studio_Code_1.35_icon.svg/3840px-Visual_Studio_Code_1.35_icon.svg.png", "download_url": "https://vscode.download.prss.microsoft.com/dbazure/download/stable/f6cfa2ea2403534de03f069bdf160d06451ed282/VSCodeUserSetup-x64-1.121.0.exe"},
                {"name": "CPU-Z", "icon_url": "https://cdn.comss.ru/logo/cpuz-icon.png?class=thumbnails", "download_url": "https://cpuz.ru/file_download/cpu-z_2.15-64bits-ru.zip"},
                {"name": "GPU-Z", "icon_url": "https://www.techpowerup.com/gpuz/icon_256.png", "download_url": "https://dl.comss.org/download/GPU-Z.2.69.0.exe"},
                {"name": "HWiNFO", "icon_url": "https://www.hwinfo.com/wp-content/themes/hwinfo/img/logo-sm.png", "download_url": "https://www.hwinfo.com/files/hwi64_846.exe"},
                {"name": "Rufus", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTI594JxCUzFPpYoo9cZNcjKVfjwlgXd3NN0Q&s", "download_url": "https://github.com/pbatard/rufus/releases/download/v4.14/rufus-4.14p.exe"},
                {"name": "Zapret", "icon_url": "https://images.steamusercontent.com/ugc/10344025336193164936/601C428D09E5F088E603B65E773DAAF722326E11/?imw=5000&imh=5000&ima=fit&impolicy=Letterbox&imcolor=%23000000&letterbox=false", "download_url": "https://github.com/Flowseal/zapret-discord-youtube/releases/download/1.9.8c/zapret-discord-youtube-1.9.8c.zip"},
                {"name": "uTorrent", "icon_url": "https://img.utdstc.com/icon/78c/99b/78c99ba5fadce0b1dea7f3b15e44020394f16daeb0cb5f702a73c551444a467e:600", "download_url": "https://www.utorrent.com/intl/ru/desktop/compare/"},
                {"name": "CapCut", "icon_url": "https://images-eds-ssl.xboxlive.com/image?url=4rt9.lXDC4H_93laV1_eHM0OYfiFeMI2p9MWie0CvL99U4GA1gf6_kayTt_kBblFwHwo8BW8JXlqfnYxKPmmBaQDG.nPeYqpMXSUQbV6ZbBQw0AEK2M5tffn2ckf3tnhLsfDJ2dtD9BYwSBOXaeJ3FkkcRWhoIPhQc9H0JHFFjc-&format=source", "download_url": "https://www.softportal.com/getsoft-49871-capcut-4.html"},
            ],
            "Библиотеки": [
                {"name": "Visual C++", "icon_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/2/20/Visual_Studio_Icon_2026.svg/1280px-Visual_Studio_Icon_2026.svg.png", "download_url": "https://dl.comss.org/download/VC_redist-2026.x64.exe"},
                {"name": ".NET 8.0", "icon_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/7/7d/Microsoft_.NET_logo.svg/1280px-Microsoft_.NET_logo.svg.png", "download_url": "https://dotnet.microsoft.com/ru-ru/download/dotnet/thank-you/runtime-aspnetcore-8.0.27-windows-x64-installer"},
                {"name": ".NET Framework", "icon_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/7/7d/Microsoft_.NET_logo.svg/1280px-Microsoft_.NET_logo.svg.png", "download_url": "https://dotnet.microsoft.com/ru-ru/download/dotnet-framework/net481"},
                {"name": "DirectX", "icon_url": "https://cdn.comss.ru/logo/directx_2022_178.png?class=thumbnails", "download_url": "https://download.microsoft.com/download/b/a/4/ba4a7e71-2906-4b2d-a0e1-80cf16844f5f/dotNetFx45_Full_setup.exe"},
                {"name": "OpenAL", "icon_url": "https://api.nuget.org/v3-flatcontainer/openal-soft/1.16.0/icon", "download_url": "https://openal.org/downloads/oalinst.zip"},
                {"name": "JDK 24", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS0azWb7Rl8nurvcSMYBgDVjG0YDP56OGYFaA&s", "download_url": "https://download.oracle.com/java/24/archive/jdk-24.0.2_windows-x64_bin.exe"},
                {"name": "JDK 21", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS0azWb7Rl8nurvcSMYBgDVjG0YDP56OGYFaA&s", "download_url": "https://download.oracle.com/java/21/archive/jdk-21.0.10_windows-x64_bin.exe"},
            ],
            "Остальное": [
                {"name": "Simple Unlocker", "icon_url": "https://private-user-images.githubusercontent.com/31757032/450704554-7ce9765f-e28c-4689-8925-6554f6bfcb0e.png", "download_url": "https://github.com/theDesConnet/SimpleUnlocker/releases/download/v1.3.2/simpleunlocker_release.zip"},
                {"name": "PowerToys", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTzPf7vMthKRCgjSxB6bxmJTypARbHiPoFgFg&s", "download_url": "https://github.com/microsoft/PowerToys/releases/download/v0.99.1/PowerToysSetup-0.99.1-x64.exe"},
                {"name": "TranslucentTB", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTo1yTzPfkuxkhOJyRoRr1PLRKfNtH5YvC4xg&s", "download_url": "https://github.com/TranslucentTB/TranslucentTB/releases/download/2026.1/bundle.msixbundle"},
                {"name": "Lively Wallpaper", "icon_url": "https://store-images.s-microsoft.com/image/apps.4720.14416131676512756.84314783-1c86-4403-b991-2e1da8525703.bf78340f-7059-4641-8d3f-8a7f740be8c0", "download_url": "https://github.com/rocksdanister/lively/releases/download/v2.2.1.0/lively_setup_x86_full_v2210.exe"},
                {"name": "OP AutoClicker", "icon_url": "https://godmode.one/wp-content/uploads/2020/02/Auto-Clicker.png", "download_url": "https://github.com/Blur009/Blur-AutoClicker/releases/download/v3.7.0/BlurAutoClicker_3.7.0_x64-setup.exe"},
                {"name": "AHK", "icon_url": "https://raw.githubusercontent.com/Ixiko/AHK-Forum/master/images/AHK%20main%20icon.png", "download_url": "https://www.autohotkey.com/download/ahk-v2.exe"},
                {"name": "Rainmeter", "icon_url": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQpRZCg2IkYrrac51gF3-TStreOwkgyMysLNw&s", "download_url": "https://github.com/rainmeter/rainmeter/releases/download/v4.5.26.3894/Rainmeter-4.5.26.exe"},
                {"name": "ExLoader", "icon_url": "https://exloader.net/resources/favicon.png", "download_url": "https://www.dropbox.com/scl/fi/wrw1znc1sz879w4x4mxj2/ExLoader_Installer-1.exe?rlkey=nbjddow3st8fa7uf78zzrb5fl&st=a7bq4lfl&dl=1"},
                {"name": "AMD Ryzen Master", "icon_url": "https://cdn.lo4d.com/t/icon/128/amd-ryzen-master.png", "download_url": "https://dl.comss.org/download/amd_ryzen_master_3_0_2.exe"},
                {"name": "MSI Afterburner", "icon_url": "https://img.icons8.com/fluent/1200/msi-afterburner.jpg", "download_url": "https://dl.comss.org/download/MSIAfterburnerSetup466.exe"},
                {"name": "Uninstall Tool", "icon_url": "https://cdn.comss.ru/logo/uninstalltool.png?class=thumbnails", "download_url": "https://github.com/crystalidea/uninstall-tool/releases/download/v3.8.0/uninstalltool_setup.exe"},
            ]
        }

    def init_ui(self):
        self.setWindowTitle("sApps - Установщик программ")
        self.setGeometry(100, 100, 1200, 800)
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.setStyleSheet("""
            QMainWindow { background-color: #292929; }
            QWidget { background-color: #292929; color: #e0e0e0; }
            QTabWidget::pane { border: none; background-color: #333333; border-radius: 10px; }
            QTabBar::tab { background-color: #3a3a3a; color: #8a8a8a; padding: 10px 20px; margin-right: 2px; border-top-left-radius: 8px; border-top-right-radius: 8px; font-size: 13px; font-weight: bold; }
            QTabBar::tab:selected { background-color: #4a4a4a; color: #4a9eff; }
            QTabBar::tab:hover:!selected { background-color: #404040; color: #b0b0b0; }
            QScrollArea { border: none; background-color: #333333; }
            QScrollBar:vertical { background-color: #333333; width: 12px; margin: 0px; }
            QScrollBar::handle:vertical { background-color: #4a4a4a; min-height: 20px; border-radius: 6px; }
            QScrollBar::handle:vertical:hover { background-color: #5a5a5a; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
            QScrollBar:horizontal { height: 0px; }
            QPushButton { background-color: #4a9eff; color: white; border: none; padding: 12px; border-radius: 8px; font-size: 14px; font-weight: bold; min-height: 20px; }
            QPushButton:hover:enabled { background-color: #3a8eef; }
            QPushButton:pressed:enabled { background-color: #2a7edf; }
            QPushButton:disabled { background-color: #5a5a5a; color: #8a8a8a; }
            QProgressBar { border: 2px solid #5a5a5a; border-radius: 6px; text-align: center; background-color: #3a3a3a; color: #e0e0e0; min-height: 22px; font-size: 10px; }
            QProgressBar::chunk { background-color: #4a9eff; border-radius: 4px; }
            QLabel { color: #e0e0e0; background: transparent; }
        """)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(25, 20, 25, 20)
        header_layout = QHBoxLayout()
        icon_label = QLabel()
        if os.path.exists(icon_path):
            pixmap = QPixmap(icon_path)
            scaled_pixmap = pixmap.scaled(40, 40, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            icon_label.setPixmap(scaled_pixmap)
            icon_label.setFixedSize(40, 40)
        header_layout.addWidget(icon_label)
        title_label = QLabel("sApps")
        title_label.setFont(QFont("Segoe UI", 22, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #e0e0e0; padding-left: 12px;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)
        self.tab_widget = QTabWidget()
        self.tabs = {}
        self.grid_layouts = {}
        programs = self.get_programs()
        for category_name in programs.keys():
            tab = QWidget()
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_area.setFrameShape(QFrame.Shape.NoFrame)
            scroll_content = QWidget()
            scroll_content.setStyleSheet("background-color: #333333;")
            grid_layout = QGridLayout(scroll_content)
            grid_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
            grid_layout.setSpacing(8)
            grid_layout.setContentsMargins(12, 12, 12, 12)
            scroll_area.setWidget(scroll_content)
            tab_layout = QVBoxLayout(tab)
            tab_layout.setContentsMargins(0, 0, 0, 0)
            tab_layout.addWidget(scroll_area)
            self.tab_widget.addTab(tab, category_name)
            self.tabs[category_name] = tab
            self.grid_layouts[category_name] = grid_layout
        main_layout.addWidget(self.tab_widget)
        bottom_panel = QWidget()
        bottom_layout = QVBoxLayout(bottom_panel)
        bottom_layout.setContentsMargins(0, 8, 0, 0)
        bottom_layout.setSpacing(10)
        self.download_button = QPushButton("📥 Скачать выбранное")
        self.download_button.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        self.download_button.clicked.connect(self.start_download)
        bottom_layout.addWidget(self.download_button)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        bottom_layout.addWidget(self.progress_bar)
        self.current_file_label = QLabel("")
        self.current_file_label.setVisible(False)
        self.current_file_label.setFont(QFont("Segoe UI", 8))
        self.current_file_label.setStyleSheet("color: #7a7a7a;")
        self.current_file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        bottom_layout.addWidget(self.current_file_label)
        main_layout.addWidget(bottom_panel)

    def clear_grid_layout(self, grid_layout):
        while grid_layout.count():
            item = grid_layout.takeAt(0)
            if item.widget():
                widget = item.widget()
                widget.setParent(None)
                widget.deleteLater()

    def get_direct_icon_url(self, url):
        if "upload.wikimedia.org" in url and "/thumb/" in url:
            parts = url.split("/")
            thumb_index = parts.index("thumb")
            del parts[thumb_index]
            del parts[-1]
            url = "/".join(parts)
        return url

    def load_icon(self, url, icon_label, app_name):
        cache_path = os.path.join(self.icon_dir, f"{app_name.replace(' ', '_')}.png")
        if os.path.exists(cache_path):
            pixmap = QPixmap(cache_path)
            if not pixmap.isNull():
                scaled_pixmap = pixmap.scaled(28, 28, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                icon_label.setPixmap(scaled_pixmap)
                return True
        try:
            session = requests.Session()
            session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
                'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
                'Referer': 'https://www.google.com/',
            })
            response = session.get(url, timeout=10, allow_redirects=True)
            if response.status_code == 200:
                pixmap = QPixmap()
                pixmap.loadFromData(response.content)
                if not pixmap.isNull():
                    try:
                        pixmap.save(cache_path, "PNG")
                    except:
                        pass
                    scaled_pixmap = pixmap.scaled(28, 28, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                    icon_label.setPixmap(scaled_pixmap)
                    return True
        except:
            pass
        icon_label.setText("📦")
        icon_label.setFont(QFont("Segoe UI", 16))
        return False

    def setup_programs(self):
        programs = self.get_programs()
        for grid_layout in self.grid_layouts.values():
            self.clear_grid_layout(grid_layout)
        self.program_cards.clear()
        for category_name, category_programs in programs.items():
            grid_layout = self.grid_layouts[category_name]
            row = 0
            col = 0
            max_cols = 3
            for i, program in enumerate(category_programs):
                card = ProgramCard()
                card.setObjectName(f"card_{category_name}_{i}")
                card.is_downloaded = self.is_app_downloaded(program["name"])
                card_style = """
                    ProgramCard {
                        background-color: #383838;
                        border-radius: 8px;
                        border: 1px solid #4a4a4a;
                    }
                """
                if not card.is_downloaded:
                    card_style += """
                        ProgramCard:hover {
                            background-color: #404040;
                            border-color: #5a9eff;
                        }
                    """
                if card.is_downloaded:
                    card.setToolTip(f"Двойной клик чтобы запустить {program['name']}")
                card.setStyleSheet(card_style)
                card_layout = QHBoxLayout(card)
                card_layout.setContentsMargins(10, 8, 10, 8)
                card_layout.setSpacing(8)
                card_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
                icon_label = QLabel()
                icon_label.setFixedSize(36, 36)
                icon_label.setStyleSheet("""
                    background-color: #4a4a4a;
                    border-radius: 6px;
                    border: 1px solid #5a5a5a;
                """)
                icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                icon_url = self.get_direct_icon_url(program["icon_url"])
                self.load_icon(icon_url, icon_label, program["name"])
                name_label = QLabel(program["name"])
                name_label.setFont(QFont("Segoe UI", 10))
                name_label.setStyleSheet("color: #e0e0e0; background: transparent; border: none;")
                name_label.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
                checkbox = CustomCheckBox()
                if card.is_downloaded:
                    checkbox.setChecked(True)
                    checkbox.setEnabled(False)
                    name_label.setStyleSheet("color: #6a6a6a; background: transparent; border: none;")
                card_layout.addWidget(icon_label)
                card_layout.addWidget(name_label)
                card_layout.addStretch()
                card_layout.addWidget(checkbox)
                card.program_data = program
                card.checkbox = checkbox
                card.name_label = name_label
                card.clicked.connect(lambda checked=False, cb=checkbox, c=card:
                                     self.toggle_card(cb, c))
                card.doubleClicked.connect(lambda app_name=program["name"]:
                                           self.launch_app(app_name))
                grid_layout.addWidget(card, row, col)
                self.program_cards.append(card)
                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1
            spacer = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
            grid_layout.addItem(spacer, row + 1, 0)

    def toggle_card(self, checkbox, card):
        if not card.is_downloaded and checkbox.enabled:
            checkbox.setChecked(not checkbox.isChecked())
            if checkbox.isChecked():
                card.setStyleSheet("""
                    ProgramCard {
                        background-color: #3a4a5a;
                        border-radius: 8px;
                        border: 1px solid #5a9eff;
                    }
                """)
            else:
                card.setStyleSheet("""
                    ProgramCard {
                        background-color: #383838;
                        border-radius: 8px;
                        border: 1px solid #4a4a4a;
                    }
                    ProgramCard:hover {
                        background-color: #404040;
                        border-color: #5a9eff;
                    }
                """)

    def start_download(self):
        if self.is_downloading:
            return
        selected_cards = [card for card in self.program_cards
                          if card.checkbox.isChecked() and not card.is_downloaded]
        if not selected_cards:
            QMessageBox.information(self, "Информация", "Выберите программы для скачивания")
            return
        self.is_downloading = True
        self.download_button.setEnabled(False)
        self.download_button.setText("⏳ Загрузка...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.current_file_label.setVisible(True)
        self.download_queue = selected_cards.copy()
        self.download_next()

    def download_next(self):
        if not self.download_queue:
            self.is_downloading = False
            self.download_button.setEnabled(True)
            self.download_button.setText("📥 Скачать выбранное")
            self.progress_bar.setVisible(False)
            self.current_file_label.setVisible(False)
            QTimer.singleShot(100, self.setup_programs)
            return
        card = self.download_queue.pop(0)
        program = card.program_data
        filename = program["download_url"].split("/")[-1]
        if "?" in filename:
            filename = filename.split("?")[0]
        if not filename or "." not in filename:
            filename = f"{program['name'].replace(' ', '_')}.exe"
        save_path = os.path.join(self.apps_dir, filename)
        if self.is_app_downloaded(program["name"]):
            print(f"Приложение уже скачано: {program['name']}")
            self.current_file_label.setText(f"✓ {program['name']} уже существует")
            self.download_next()
            return
        if os.path.exists(save_path):
            try:
                os.remove(save_path)
            except:
                pass
        print(f"Скачивание {program['name']} из {program['download_url']}")
        print(f"Сохранение в: {save_path}")
        self.download_thread = DownloadThread(program["download_url"], save_path, program["name"])
        self.download_thread.progress.connect(self.progress_bar.setValue)
        self.download_thread.current_file.connect(self.current_file_label.setText)
        self.download_thread.finished.connect(self.on_download_finished)
        self.download_thread.error.connect(self.on_download_error)
        self.download_thread.start()

    def on_download_finished(self, save_path, app_name):
        print(f"Файл сохранен: {save_path}")
        self.save_downloaded_app(app_name, save_path)
        start_menu_apps = ["YouTube Music", "Rufus"]
        if app_name in start_menu_apps:
            self.create_start_menu_shortcut(app_name, save_path)
        QTimer.singleShot(200, lambda: self.show_success_dialog(app_name, save_path))

    def show_success_dialog(self, app_name, save_path):
        dialog = SuccessDialog(app_name, save_path, self)
        dialog.exec()
        QTimer.singleShot(100, self.setup_programs)
        QTimer.singleShot(200, self.download_next)

    def on_download_error(self, error_message):
        print(f"Ошибка скачивания: {error_message}")
        self.current_file_label.setText(f"✗ Ошибка: {error_message[:80]}...")
        QTimer.singleShot(200, lambda: self.show_error_dialog(error_message))

    def show_error_dialog(self, error_message):
        error_dialog = QMessageBox()
        error_dialog.setIcon(QMessageBox.Icon.Critical)
        error_dialog.setWindowTitle("Ошибка скачивания")
        error_dialog.setText("Не удалось скачать файл")
        error_dialog.setInformativeText(f"{error_message[:200]}\n\nПроверьте подключение к интернету и попробуйте снова.")
        error_dialog.setStyleSheet("""
            QMessageBox {
                background-color: #2a2a2a;
                color: #e0e0e0;
            }
            QPushButton {
                background-color: #4a9eff;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 6px;
                font-size: 13px;
                min-width: 100px;
            }
        """)
        error_dialog.exec()
        QTimer.singleShot(100, self.download_next)

    def closeEvent(self, event):
        if hasattr(self, 'download_thread') and self.download_thread.isRunning():
            self.download_thread.quit()
            self.download_thread.wait(1000)
        event.accept()

def main():
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        if not is_admin:
            print("программа запущена без прав администратора")
    except:
        pass
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(41, 41, 41))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(224, 224, 224))
    palette.setColor(QPalette.ColorRole.Base, QColor(41, 41, 41))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(51, 51, 51))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(41, 41, 41))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(224, 224, 224))
    palette.setColor(QPalette.ColorRole.Text, QColor(224, 224, 224))
    palette.setColor(QPalette.ColorRole.Button, QColor(41, 41, 41))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(224, 224, 224))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(74, 158, 255))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)
    window = DownloaderApp()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()