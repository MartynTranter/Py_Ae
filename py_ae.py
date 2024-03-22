"""
# Python After Effects (PyAe)

* Description

    A simple PyQt5 Gui for managing After Effects projects and project assets.

* Update History

    `2024-03-17` - Init
"""

import getpass
import os
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import psutil
import pygetwindow
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings, QProcess, QSize
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QFileDialog, QHBoxLayout, QInputDialog, QLineEdit, QLabel, QFormLayout, QDialog, QDialogButtonBox)


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        """
        Initializes the settings dialog window.

        Args:
            parent (QWidget, optional): The parent widget. Defaults to None.
        """
        super(SettingsDialog, self).__init__(parent)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self.setWindowTitle("Settings")
        self.layout = QFormLayout(self)

        # Setup UI components for After Effects path
        self.aePathEdit = QLineEdit(self)
        self.aePathBrowseButton = QPushButton("Browse")
        aePathLayout = QHBoxLayout()
        aePathLayout.addWidget(self.aePathEdit)
        aePathLayout.addWidget(self.aePathBrowseButton)
        self.layout.addRow(QLabel("After Effects Path:"), aePathLayout)

        # Setup UI components for assets folder selection
        self.assetsFolderEdit = QLineEdit(self)
        self.assetsFolderBrowseButton = QPushButton("Browse")
        assetsFolderLayout = QHBoxLayout()
        assetsFolderLayout.addWidget(self.assetsFolderEdit)
        assetsFolderLayout.addWidget(self.assetsFolderBrowseButton)
        self.layout.addRow(QLabel("Assets Folder:"), assetsFolderLayout)

        # Setup UI components for projects save folder selection
        self.projectsFolderEdit = QLineEdit(self)
        self.projectsFolderBrowseButton = QPushButton("Browse")
        projectsFolderLayout = QHBoxLayout()
        projectsFolderLayout.addWidget(self.projectsFolderEdit)
        projectsFolderLayout.addWidget(self.projectsFolderBrowseButton)
        self.layout.addRow(QLabel("Projects Folder:"), projectsFolderLayout)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.layout.addRow(self.buttons)

        # Connect signals and slots for dialog interactions
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.aePathBrowseButton.clicked.connect(self.browseForAePath)
        self.assetsFolderBrowseButton.clicked.connect(self.browseForAssetsFolder)
        self.projectsFolderBrowseButton.clicked.connect(self.browseForProjectsFolder)

    def getSettings(self):
        return self.aePathEdit.text(), self.assetsFolderEdit.text(), self.projectsFolderEdit.text()

    def browseForAePath(self):
        filePath, _ = QFileDialog.getOpenFileName(self, "Select After Effects Executable", "",
                                                  "Executable Files (*.exe)")
        if filePath:
            self.aePathEdit.setText(filePath)

    def browseForAssetsFolder(self):
        folderPath = QFileDialog.getExistingDirectory(self, "Select Assets Folder")
        if folderPath:
            self.assetsFolderEdit.setText(folderPath)

    def browseForProjectsFolder(self):
        folderPath = QFileDialog.getExistingDirectory(self, "Select Projects Save Folder")
        if folderPath:
            self.projectsFolderEdit.setText(folderPath)

    def accept(self):
        settings = QSettings("YourOrganization", "AfterEffectsPipeline")
        settings.setValue("assetsFolder", self.assetsFolderEdit.text())
        settings.setValue("aePath", self.aePathEdit.text())
        settings.setValue("projectsFolder", self.projectsFolderEdit.text())

        super().accept()

    def clear_projects_table(self):
        """
        Clears all rows from the projects table.
        """
        while self.projects_table.rowCount() > 0:
            self.projects_table.removeRow(0)


class AfterEffectsPipeline(QWidget):
    """
    A graphical user interface for managing Adobe After Effects projects.

    This application allows users to open, create, and manage Adobe After Effects projects.
    It provides functionalities such as searching for projects, opening existing projects, creating new projects,
    importing files into After Effects, and running After Effects. The interface is built using PyQt5.

    Attributes:
        last_project_directory (Path): The directory of the last accessed project.
        last_project_files (list[Path]): A list of paths for the last accessed project files.
        imported_assets (list[Path]): A list of paths for the imported assets.
    """

    def __init__(self):
        super().__init__()

        # Initialize the user interface
        # Set the window attributes
        self.setWindowTitle("After Effects Browser")
        self.setWindowIcon(QIcon('images/PyAE_icon.ico'))
        self.resize(1000, 1000)
        self.setMinimumSize(600, 500)

        settings = QSettings("YourOrganization", "AfterEffectsPipeline")
        self.aePath = settings.value("aePath", "default/path/to/AfterEffects.exe")  # Adjust default path as needed
        self.assetsFolder = settings.value("assetsFolder", "default/path/to/assets")  # Adjust default path as needed

        # Initialize the last project directory and list variables
        self.last_project_directory = None
        self.last_project_files: list[Path] = []
        self.cur_project_files: list[Path] = []
        self.process = QProcess(self)
        self.process.finished.connect(
            self.on_process_finished)

        self.set_qss()
        self.create_search_bar()
        self.create_asset_search_bar()
        self.create_widgets()
        self.create_layout()
        self.create_connections()
        self._load_project_settings()
        self._load_asset_settings()
        self._load_window_settings()

    def set_qss(self):
        qss_path = Path(Path(__file__).parent, 'Combinear.qss')
        with open(qss_path, 'r') as f:
            stylesheet = f.read()
            self.setStyleSheet(stylesheet)

    def show_settings_dialog(self):
        """
        Displays a settings dialog window to allow the user to configure various settings.

        Reads previously saved settings from QSettings and pre-populates the dialog fields with them.
        Upon closing the dialog, saves the updated settings if the user accepts the changes.
        """
        dialog = SettingsDialog(self)
        settings = QSettings("YourOrganization", "AfterEffectsPipeline")

        # Pre-populate dialog fields with saved settings
        dialog.aePathEdit.setText(settings.value("aePath", ""))
        dialog.assetsFolderEdit.setText(settings.value("assetsFolder", ""))
        dialog.projectsFolderEdit.setText(settings.value("projectsFolder", ""))  # New line to handle project folder

        # Show the dialog and save settings if OK is pressed
        if dialog.exec_() == QDialog.Accepted:
            aePath, assetsFolder, projectsFolder = dialog.getSettings()  # Updated to capture the projectsFolder
            settings.setValue("aePath", aePath)
            settings.setValue("assetsFolder", assetsFolder)
            settings.setValue("projectsFolder", projectsFolder)  # New line to save the project folder setting


    @staticmethod
    def activate_after_effects_window():
        """
        Activates and positions the After Effects window.

        This method attempts to find the After Effects window either by an exact title match or a partial match.
        Once the window is found, it waits for a brief moment to ensure the application is ready before adjusting its position.

        Note:
            This method depends on the pygetwindow library for window manipulation.

        Raises:
            Exception: If there's an error while managing the After Effects window.
        """
        # Adjust the title to match the exact title of the After Effects window, including the version year if necessary
        window_title_exact = "Adobe After Effects 2024"
        ae_windows = pygetwindow.getWindowsWithTitle(window_title_exact)
        if not ae_windows:  # If not found, try a partial match
            ae_windows = [win for win in pygetwindow.getAllWindows() if "Adobe After Effects" in win.title]

        if ae_windows:
            try:
                # If the window was found, proceed to adjust the position
                # Wait a moment before adjusting window position to ensure the application is ready
                time.sleep(2)

                # Now using the exact or partial match for the window title
                x, y, width, height = 100, 100, 1280, 720  # Example values
                AfterEffectsPipeline.set_window_position(window_title_exact, x, y, width, height)
            except Exception as e:
                print(f"[1] Error managing After Effects window: {e}")

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Construction
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    def create_search_bar(self):
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search Projects")
        self.search_bar.textChanged.connect(self.filter_projects)

    def create_asset_search_bar(self):
        self.asset_search_bar = QLineEdit()
        self.asset_search_bar.setPlaceholderText("Search Assets")
        self.asset_search_bar.textChanged.connect(self.filter_assets)

    def create_widgets(self):
        """
        Creates UI widgets for the application.

        Initializes buttons for running After Effects, opening a project, creating a new project, and importing a file.
        Also, initializes the table widget for displaying last opened projects with appropriate headers and setup.
        """
        self.run_after_effects_button = QPushButton()
        self.run_after_effects_button.setText("Open After Effects")
        self.run_after_effects_button.setIcon(QIcon("images/After_effects_icon.svg"))
        self.run_after_effects_button.setIconSize(QSize(20, 20))
        self.run_after_effects_button.setFixedSize(80, 80)

        self.open_project_button = QPushButton()
        self.open_project_button.setText("Open Project")
        self.open_project_button.setIcon(QIcon("images/book-open.svg"))  # Set the color here
        self.open_project_button.setIconSize(QSize(20, 20))  # Set icon size (optional)
        self.open_project_button.setFixedSize(80, 80)

        self.new_project_button = QPushButton()
        self.new_project_button.setText("New Project")
        self.new_project_button.setIcon(QIcon("images/file-plus.svg"))  # Set the color here
        self.new_project_button.setIconSize(QSize(20, 20))  # Set icon size (optional)
        self.new_project_button.setFixedSize(80, 80)

        self.import_file_button = QPushButton()
        self.import_file_button.setText("Import Asset")
        self.import_file_button.setIcon(QIcon("images/file-plus.svg"))
        self.import_file_button.setIconSize(QSize(20, 20))  # Set icon size (optional)
        self.import_file_button.setFixedSize(80, 80)

        self.settings_button = QPushButton()
        self.settings_button.setText("Settings")
        self.settings_button.setIcon(QIcon('images/settings.svg'))  # Ensure you have an appropriate icon
        self.settings_button.setIconSize(QSize(20, 20))
        self.settings_button.setFixedSize(200, 40)

        self.run_after_effects_button.setFixedSize(200, 40)
        self.open_project_button.setFixedSize(200, 40)
        self.new_project_button.setFixedSize(200, 40)
        self.import_file_button.setFixedSize(200, 40)

        # Create a table widget to display last project files
        self.projects_table = QTableWidget()
        self.projects_table.setColumnCount(5)  # Increase column count for buttons
        self.projects_table.setHorizontalHeaderLabels(
            ['Project Name', 'File Path', 'Last Modified', 'Created By', 'Actions'])
        self.projects_table.horizontalHeader().setStretchLastSection(True)
        self.projects_table.setSortingEnabled(True)
        self.projects_table.setSelectionBehavior(QTableWidget.SelectRows)

        # Create a table widget to display imported assets
        self.imported_assets_table = QTableWidget()
        self.imported_assets_table.setColumnCount(5)  # Adjusted from 4 to 5
        self.imported_assets_table.setHorizontalHeaderLabels(
            ['Asset Name', 'File Path', 'Last Modified', 'Imported By', 'Actions'])
        self.imported_assets_table.horizontalHeader().setStretchLastSection(True)
        self.imported_assets_table.setSortingEnabled(True)
        self.imported_assets_table.setSelectionBehavior(QTableWidget.SelectRows)

    def create_layout(self):
        """
        Arranges UI widgets into the main application layout.

        Sets up the application layout by organizing buttons in a horizontal layout at the top, project and asset tables
        in a vertical layout below, integrating the search bar, and aligning everything within the main window.
        """
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.run_after_effects_button)
        button_layout.addWidget(self.open_project_button)
        button_layout.addWidget(self.new_project_button)
        button_layout.addWidget(self.import_file_button)
        button_layout.addWidget(self.settings_button)
        button_layout.addStretch()

        # Arrange the last projects table and imported assets table in separate layouts
        projects_layout = QVBoxLayout()

        # Add a label above the last projects table
        saved_projects_label = QLabel("Saved Projects")
        saved_projects_label.setAlignment(Qt.AlignLeft)  # Align the text to the left
        projects_layout.addWidget(saved_projects_label)
        projects_layout.addWidget(self.search_bar)
        projects_layout.addWidget(self.projects_table)

        # Add a label above the imported assets table
        imported_assets_label = QLabel("Imported Assets")
        imported_assets_label.setAlignment(Qt.AlignLeft)  # Align the text to the left
        projects_layout.addWidget(imported_assets_label)
        projects_layout.addWidget(self.asset_search_bar)
        projects_layout.addWidget(self.imported_assets_table)

        # Create the main layout by combining the button layout and table layout
        main_layout = QVBoxLayout(self)
        main_layout.addLayout(button_layout)
        main_layout.addLayout(projects_layout)

    def create_connections(self):
        """
        Connects UI signals to their corresponding slots or methods.

        Establishes connections between UI components and their handling methods, ensuring the application
        responds appropriately to user actions.
        """
        self.run_after_effects_button.clicked.connect(self.run_after_effects_connection)
        self.open_project_button.clicked.connect(self.open_after_effects_project_connection)
        self.import_file_button.clicked.connect(self.import_file_connection)
        self.new_project_button.clicked.connect(self.create_new_project_connection)
        self.settings_button.clicked.connect(self.show_settings_dialog)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Save Settings
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    def closeEvent(self, event):
        self._save_project_settings()
        self._save_asset_settings()
        self._save_window_settings()
        super(AfterEffectsPipeline, self).closeEvent(event)

    def _save_window_settings(self):
        save_path = Path(os.getenv("USERPROFILE"), f'{self.windowTitle()}Settings.ini')
        save_settings_obj = QSettings(save_path.as_posix(), QSettings.IniFormat)
        save_settings_obj.setValue('window_geo', self.saveGeometry())

    def _load_window_settings(self):
        save_path = Path(os.getenv("USERPROFILE"), f'{self.windowTitle()}Settings.ini')
        save_settings_obj = QSettings(save_path.as_posix(), QSettings.IniFormat)
        value = save_settings_obj.value('window_geo')
        if value:
            self.restoreGeometry(value)

    def _save_project_settings(self):
        save_path = Path(os.getenv("USERPROFILE"), f'{self.windowTitle()}ProjectSettings.ini')
        save_settings_obj = QSettings(save_path.as_posix(), QSettings.IniFormat)

        save_data = []
        table_data = self._collect_table_data(self.projects_table)
        for data in table_data:
            cur_data = []
            for _, v in data.items():
                cur_data.append(v)
            save_data.append(cur_data)

        save_settings_obj.setValue("project_table", save_data)

    def _load_project_settings(self):
        save_path = Path(os.getenv("USERPROFILE"), f'{self.windowTitle()}ProjectSettings.ini')
        save_settings_obj = QSettings(save_path.as_posix(), QSettings.IniFormat)
        save_data = save_settings_obj.value('project_table')
        if not save_data:
            return
        for i in save_data:
            if not Path(i[1]).exists():
                continue
            table_data = self._create_blank_project_data()
            table_data['Project Name'] = i[0]
            table_data['File Path'] = i[1]
            table_data['Last Modified'] = i[2]
            table_data['Created By'] = i[3]
            self.add_project_row(table_data)

    def _save_asset_settings(self):
        save_path = Path(os.getenv("USERPROFILE"), f'{self.windowTitle()}AssetsSettings.ini')
        save_settings_obj = QSettings(save_path.as_posix(), QSettings.IniFormat)

        save_data = []
        table_data = self._collect_asset_table_data(self.imported_assets_table)
        for data in table_data:
            cur_data = []
            for _, v in data.items():
                cur_data.append(v)
            save_data.append(cur_data)

        save_settings_obj.setValue("imported_assets_table", save_data)

    def _load_asset_settings(self):
        save_path = Path(os.getenv("USERPROFILE"), f'{self.windowTitle()}AssetsSettings.ini')
        save_settings_obj = QSettings(save_path.as_posix(), QSettings.IniFormat)
        save_data = save_settings_obj.value('imported_assets_table')
        if not save_data:
            return
        for i in save_data:
            if not Path(i[1]).exists():
                continue
            table_data = self._create_asset_data()
            table_data['Asset Name'] = i[0]
            table_data['File Path'] = i[1]
            table_data['Last Modified'] = i[2]
            table_data['Imported By'] = i[3]
            self.add_asset_row(table_data)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Front end functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    def filter_projects(self, sub_string: str):
        """
        Filters displayed projects based on search input.

        Args:
            sub_string (str): Search query entered by the user.
        """
        print(f'need to rewrite - substring: {sub_string}')

    def filter_assets(self):
        """
        Filters the displayed assets based on the search text entered by the user.

        Retrieves the search text from the asset search bar and iterates through each row in the
        imported assets table. For each asset name in the first column, checks if the search text
        is present. If found, the corresponding row is displayed; otherwise, it's hidden.

        Note:
            This method assumes that the imported_assets_table is already populated with asset data.
        """
        search_text = self.asset_search_bar.text().lower()
        for row in range(self.imported_assets_table.rowCount()):
            item = self.imported_assets_table.item(row, 0)  # Get the item in the first column (Asset Name)
            if item is not None:
                asset_name = item.text().lower()
                # Check if the search text is present in the asset name
                if search_text in asset_name:
                    self.imported_assets_table.setRowHidden(row, False)
                else:
                    self.imported_assets_table.setRowHidden(row, True)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Input Connections
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    def open_project_connection(self, project_path: Path):
        """
        Opens an After Effects project.

        Opens the specified After Effects project or the project selected in the last projects table if no path is provided.
        It handles the subprocess call to After Effects with the project path.

        Args:
            project_path (str, optional): The file path to the After Effects project to be opened. If None, the project
            selected in the last projects table will be opened. Defaults to None.
        """
        if project_path:
            aerender_path = Path("C:/Program Files/Adobe/Adobe After Effects 2024/Support Files/AfterFX.exe")
            command = [aerender_path.as_posix(), project_path]

            # Execute the After Effects command asynchronously
            self.run_after_effects_command_async(command)
        else:
            selected_row = self.projects_table.currentRow()
            if selected_row == -1:
                self.show_error_message("No project selected", "Please select a project to open.")
                return
            project_path_item = self.projects_table.item(selected_row, 1)
            if project_path_item:
                project_path = project_path_item.text()
                aerender_path = Path("C:/Program Files/Adobe/Adobe After Effects 2024/Support Files/AfterFX.exe")
                command = [aerender_path.as_posix(), project_path]

                # Call the asynchronous method to run After Effects command
                self.run_after_effects_command_async(command)
            else:
                self.show_error_message("Invalid project", "Selected project is invalid.")
                return

    def delete_project(self, index: int):
        """
        Deletes a project file from the filesystem and updates the UI.

        Args:
            index (int): The index of row and the related project.
        """
        # Confirm deletion with the user
        confirmation = QMessageBox.question(
            self, "Confirmation", "Are you sure you want to delete this project?",
            QMessageBox.Yes | QMessageBox.Cancel)

        if confirmation != QMessageBox.Yes:
            return

        # # Attempt to delete the project file
        project_path = Path(self.projects_table.item(index, 1).text())
        if not project_path.exists():
            QMessageBox.information(self, "File Not Found!")
            return

        os.remove(project_path)
        self.remove_row(index)

    def create_new_project_connection(self):
        # Use the user-defined projects folder
        projectsFolder = Path(QSettings("YourOrganization", "AfterEffectsPipeline").value("projectsFolder",
                                                                                          "default/path/to/projects"))
        if not projectsFolder.exists():
            os.makedirs(projectsFolder, exist_ok=True)

        try:
            project_name, ok_pressed = QInputDialog.getText(self, "New Project", "Enter the project name:")
            if ok_pressed and project_name:
                new_project_path = Path(projectsFolder, f"{project_name}.aep")
                template_project_path = Path(Path(__file__).parent, 'Blank_Project.aep')
                shutil.copyfile(template_project_path, new_project_path.as_posix())

                # Fetch the current operating system's username
                created_by = getpass.getuser()  # Use the OS username as the project creator's name

                data = self._create_blank_project_data()
                data['Project Name'] = project_name
                data['File Path'] = new_project_path.as_posix()
                data['Last Modified'] = datetime.fromtimestamp(new_project_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                data['Created By'] = getpass.getuser()
                self.add_project_row(data)
        except Exception as e:
            print(f"An error occurred: {e}")

    def run_after_effects_connection(self):
        """
        Launches After Effects without opening a project.

        Executes a subprocess call to launch Adobe After Effects. If the path to the After Effects executable is not
        provided, it displays an error message.
        """

        settings = QSettings("YourOrganization", "AfterEffectsPipeline")
        aePath = settings.value("aePath", "default/path/to/AfterEffects.exe")  # Fallback default
        if not os.path.exists(aePath):
            self.show_error_message("After Effects path not provided", "Executable not found.")
            return
        # Run After Effects with the specified command
        aerender_path = r"C:/Program Files/Adobe/Adobe After Effects 2024/Support Files/AfterFX.exe"
        if aerender_path:
            command = [aerender_path]
            try:
                # Use subprocess.Popen with start method
                subprocess.Popen(command, start_new_session=True)
            except Exception as e:
                print(f"[2] Error launching After Effects: {e}")
        else:
            self.show_error_message("After Effects path not provided", "Please enter the path to After Effects.")

    def open_after_effects_project_connection(self):
        """
        Opens a selected After Effects project file.

        Prompts the user to select an After Effects project file via a file dialog. Updates the last project directory
        and list based on the selection, saves the updated information to settings, and launches After Effects with the
        selected project.
        """
        aerender_path = Path("C:/Program Files/Adobe/Adobe After Effects 2024/Support Files/AfterFX.exe")
        if not Path(aerender_path).exists():
            self.show_error_message("After Effects path not provided", "Executable not found.")
            return

        initial_dir = str(self.last_project_directory) if self.last_project_directory else ""
        project_path, _ = QFileDialog.getOpenFileName(self, "Open After Effects Project", initial_dir,
                                                      "After Effects Project Files (*.aep);;All Files (*)")
        if project_path:
            print(f"Project selected: {project_path}")  # Diagnostic print
            self.last_project_directory = Path(project_path).parent
            # Use QProcess to start After Effects with the selected project
            self.process = QProcess(self)
            self.process.setProgram(aerender_path)
            self.process.setArguments([project_path])

            # Optional: Connect signals for process start, finish, and error handling
            self.process.started.connect(lambda: print("After Effects started."))
            self.process.finished.connect(lambda exitCode: print(f"After Effects finished with exit code {exitCode}."))
            self.process.errorOccurred.connect(lambda error: print(f"[3] Error occurred: {self.process.errorString()}"))

            self.process.start()  # Start the process
            print("Command to open After Effects has been issued.")

    @staticmethod
    def on_process_finished():
        print("After Effects command executed successfully.")

    @staticmethod
    def is_after_effects_running():
        """
        Checks if Adobe After Effects is running.

        Returns:
            bool: True if After Effects is running, False otherwise.
        """
        # The name of the After Effects executable might vary; adjust as needed
        ae_process_name = "AfterFX.exe"
        for proc in psutil.process_iter(attrs=['name']):
            if ae_process_name in proc.info['name']:
                return True
        return False

    def import_file_connection(self):
        """Imports a selected file into the current After Effects project, updating the UI and settings accordingly."""
        if not self.is_after_effects_running():
            self.show_error_message("After Effects Not Running", "Adobe After Effects must be open to import assets.")
            return

        if not self.after_effects_path.exists():
            return

        if not self.after_effects_path.exists():
            return

        file_path = Path(self.prompt_user_for_file(self.after_effects_path))
        if not file_path.exists():
            return

        # Copy the file to the assets folder and prepare for import
        destination = Path(self.assets_folder_path, file_path.name)
        shutil.copy(file_path.as_posix(), destination.as_posix())

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Generate and execute JSX script for importing
        self.execute_jsx_script(self.after_effects_path.as_posix(), destination.as_posix())
        # Update UI and settings
        self.update_imported_assets(file_path, current_time, getpass.getuser())

    @property
    def after_effects_path(self) -> Optional[Path]:
        ae_path = Path(QSettings("YourOrganization", "AfterEffectsPipeline").value("aePath", ""))
        if ae_path.exists() and ae_path.is_file():
            return ae_path

        return Path('does/not/exist/dasds')

    @property
    def assets_folder_path(self) -> Optional[Path]:
        asset_folder_path = Path(QSettings("YourOrganization", "AfterEffectsPipeline").value("assetsFolder", ""))
        if asset_folder_path and asset_folder_path.exists():
            return asset_folder_path

        return Path('does/not/exist/dasds')

    def prompt_user_for_file(self, asset_folder: Path) -> Optional[Path]:
        file_name, _ = QFileDialog.getOpenFileName(self, "Import File into After Effects", asset_folder.as_posix(),
                                                   "All Files (*)")
        return file_name if file_name else None

    def execute_jsx_script(self, ae_path, file_path):
        corrected_path = file_path.replace('//', '////')
        script_content = f"""
            var fileToImport = new File("{corrected_path}");
            if (fileToImport.exists) {{
                app.project.importFile(new ImportOptions(fileToImport));
            }} else {{
                alert('File does not exist: {corrected_path}');
            }}
        """
        script_file_path = tempfile.mktemp(suffix=".jsx")
        with open(script_file_path, 'w') as script_file:
            script_file.write(script_content)
        self.run_after_effects_command_async([ae_path, '-r', script_file_path])

    def update_imported_assets(self, file_path: Path, last_modified: str, imported_by: str):
        self.populate_imported_assets_table(file_path, last_modified, imported_by)

    def populate_imported_assets_table(self, file_path: Path, last_modified: str, imported_by: str):
        """
        Populates the imported assets table with the new asset.

        Args:
            file_path (Path): The file path of the imported asset.
            last_modified (str): Last modified timestamp of the asset.
            imported_by(str): The user who imported it.
        """
        row_count = self.imported_assets_table.rowCount()
        self.imported_assets_table.insertRow(row_count)
        self.imported_assets_table.setItem(row_count, 0, QTableWidgetItem(file_path.name))
        self.imported_assets_table.setItem(row_count, 1, QTableWidgetItem(file_path.as_posix()))
        self.imported_assets_table.setItem(row_count, 2, QTableWidgetItem(last_modified))
        self.imported_assets_table.setItem(row_count, 3, QTableWidgetItem(imported_by))

        # Create a container widget to hold both action buttons
        actions_widget = QWidget()
        actions_layout = QHBoxLayout()
        actions_layout.setContentsMargins(0, 0, 0, 0)

        # Add "Import to Current Project" button
        import_button = QPushButton("Import")
        import_button.clicked.connect(lambda _, path=file_path: self.import_to_current_project(path))
        actions_layout.addWidget(import_button)

        # Add "Delete" button
        delete_button = QPushButton("Delete")
        delete_button.clicked.connect(lambda _, row=row_count: self.delete_imported_asset(row))
        actions_layout.addWidget(delete_button)
        actions_widget.setLayout(actions_layout)

        self.imported_assets_table.setCellWidget(row_count, 4, actions_widget)
        self.imported_assets_table.setRowCount(row_count + 1)

    def import_to_current_project(self, file_path: Path):
        """
        Imports the selected file into the current After Effects project.

        Args:
            file_path (path): The file path of the asset to import into After Effects.
        """

        if not self.is_after_effects_running():
            self.show_error_message("After Effects Not Running", "Adobe After Effects must be open to perform this action.")
            return

        ae_path = self.after_effects_path
        if ae_path is None or not ae_path.exists():
            self.show_error_message("Invalid Path", "The path to After Effects is invalid.")
            return

        # Construct the JSX script for importing the asset
        jsx_script = f"""
            var myFile = new File("{file_path.as_posix()}");
            app.project.importFile(new ImportOptions(myFile));
            """

        # Save the JSX script to a temporary file
        temp_jsx_path = Path(tempfile.gettempdir(), "tempImportScript.jsx")
        with open(temp_jsx_path, "w") as jsx_file:
            jsx_file.write(jsx_script)

        # Run the script in After Effects
        command = [ae_path.as_posix(), "-r", temp_jsx_path.as_posix()]
        subprocess.run(command)

        # Optionally, delete the temp script after execution
        temp_jsx_path.unlink(missing_ok=True)

    def delete_imported_asset(self, row: int):
        confirmation = QMessageBox.question(self, "Confirmation", "Are you sure?", QMessageBox.Yes | QMessageBox.No)
        if confirmation == QMessageBox.Yes:
            asset_path = Path(self.imported_assets_table.itemAt(row, 1).text())
            if not asset_path.exists():
                self.imported_assets_table.removeRow(row)
                return

            os.remove(asset_path)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Back-end Functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    @staticmethod
    def _create_blank_project_data() -> dict:
        """['Project Name', 'File Path', 'Last Modified', 'Created By', 'Actions']"""
        data = {
            'Project Name': '',
            'File Path': '',
            'Last Modified': '',
            'Created By': ''
        }
        return data

    @staticmethod
    def _create_asset_data() -> dict:
        """['Asset Name', 'File Path', 'Last Modified', 'Imported By', 'Actions']"""
        data = {
            'Asset Name': '',
            'File Path': '',
            'Last Modified': '',
            'Imported By': ''
        }
        return data

    def remove_row(self, index: int):
        """
        Remove a row from the last projects table.

        This method removes a row at the specified index from the last projects table.
        It also updates the project list and saves the changes to the application settings.

        Args:
            index (int): The index of the row to be removed.
        """
        count = self.projects_table.rowCount()
        if count > index:
            self.projects_table.removeRow(index)

    def add_project_row(self, row_data: dict):
        new_row_index = self.projects_table.rowCount()
        self.projects_table.insertRow(new_row_index)
        for i, v in enumerate(row_data.values()):
            self.projects_table.setItem(new_row_index, i, QTableWidgetItem(v))
        action_widget = self._create_project_widget(new_row_index, row_data['File Path'])
        self.projects_table.setCellWidget(new_row_index, 4, action_widget)  # Hard coded column value

    def add_asset_row(self, row_data: dict):
        new_row_index = self.imported_assets_table.rowCount()
        self.imported_assets_table.insertRow(new_row_index)
        for i, v in enumerate(row_data.values()):
            self.imported_assets_table.setItem(new_row_index, i, QTableWidgetItem(v))
        action_widget = self._create_asset_widget(new_row_index)
        self.imported_assets_table.setCellWidget(new_row_index, 4, action_widget)  # Hard coded column value

    def run_after_effects_command_async(self, command: list[str]):
        """
        Runs an After Effects command asynchronously using a QThread.

        Args:
            command (list[str]): The After Effects command to execute asynchronously.
        """
        self.thread = AfterEffectsThread(command, self)
        self.thread.finished.connect(self.thread_finished)
        self.thread.start()

    def _collect_table_data(self, table: QTableWidget) -> list[dict]:
        table_data = []
        for row in range(table.rowCount()):
            cur_proj_data = self._create_blank_project_data()
            for col in range(table.columnCount()):
                key = table.horizontalHeaderItem(col).text()
                try:
                    cur_proj_data[key] = table.item(row, col).text()
                except AttributeError:
                    pass
            table_data.append(cur_proj_data)

        return table_data

    def _collect_asset_table_data(self, table: QTableWidget) -> list[dict]:
        table_data = []
        for row in range(table.rowCount()):
            cur_asset_data = self._create_asset_data()
            for col in range(table.columnCount()):
                key = table.horizontalHeaderItem(col).text()
                try:
                    cur_asset_data[key] = table.item(row, col).text()
                except AttributeError:
                    pass
            table_data.append(cur_asset_data)
        return table_data

    def _create_project_widget(self, row_index: int, project_path: Path) -> QWidget:
        """
        Container widget for project table

        Args:
            row_index(int): The row index from the table widget for button connections.
            project_path(Path): Which project path to open.

        Returns:
            QtWidgets.QWidget: The widget containing the open and delete project buttons.
        """
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        open_button = QPushButton('Open')
        delete_button = QPushButton('Delete')

        open_button.clicked.connect(lambda: self.open_project_connection(project_path))
        delete_button.clicked.connect(lambda: self.delete_project(row_index))

        actions_layout.addWidget(open_button)
        actions_layout.addWidget(delete_button)

        return actions_widget

    def _create_asset_widget(self, row_index: int) -> QWidget:
        """
        Container widget for asset table.

        Args:
            row_index(int): The row index from the table widget for button connections.

        Returns:
            QtWidgets.QWidget: The widget containing the open and delete project buttons.
        """
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(0, 0, 0, 0)
        open_button = QPushButton('Import to Last Opened')
        delete_button = QPushButton('Delete')

        file_path = Path(self.imported_assets_table.itemAt(row_index, 1).text())

        open_button.clicked.connect(lambda: self.import_to_current_project(file_path))
        delete_button.clicked.connect(lambda: self.delete_imported_asset(row_index))

        actions_layout.addWidget(open_button)
        actions_layout.addWidget(delete_button)

        return actions_widget

    @staticmethod
    def thread_finished(success: bool):
        if success:
            print("After Effects command executed successfully.")
        else:
            print("After Effects command failed.")

    @staticmethod
    def show_error_message(title, message):
        error_message = QMessageBox()
        error_message.setIcon(QMessageBox.Critical)
        error_message.setWindowTitle(title)
        error_message.setText(message)
        error_message.exec_()

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Saved Settings Handling 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


class AfterEffectsThread(QThread):
    finished = pyqtSignal(bool)

    def __init__(self, command: list[str], parent: QWidget):
        super().__init__()
        self.command = ' '.join(command)
        self.parent = parent

    def run(self):
        try:
            subprocess.run(self.command, check=True)
            self.finished.emit(True)
        except subprocess.CalledProcessError as e:
            print(f"[5] Error during After Effects execution: {e}")
            self.finished.emit(False)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AfterEffectsPipeline()
    window.show()

    app.exec_()
