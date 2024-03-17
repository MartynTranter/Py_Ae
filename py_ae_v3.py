# Standard library imports
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
import getpass

# Third-party library imports
import psutil
import pygetwindow
import win32gui
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings, QProcess, QSize
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QFileDialog, QHBoxLayout, QInputDialog, QLineEdit, QLabel, QFormLayout, QDialog, QDialogButtonBox)



def colorize_svg(svg_path, color=Qt.white):
    pixmap = QPixmap(svg_path)
    return QIcon(pixmap)


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
        filePath, _ = QFileDialog.getOpenFileName(self, "Select After Effects Executable", "", "Executable Files (*.exe)")
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
        self.imported_assets: list[dict] = []
        qss_path = Path(Path(__file__).parent, 'Combinear.qss')
        with open(qss_path, 'r') as f:
            stylesheet = f.read()
            self.setStyleSheet(stylesheet)
            self.process = QProcess(self)  # Add this line to initialize QProcess
            self.process.finished.connect(self.onProcessFinished)  # Optional: Connect to a method to handle process completion

        self.create_search_bar()
        self.create_asset_search_bar()
        self.create_widgets()
        self.create_layout()
        self.create_connections()
        self.load_settings()

    def showSettingsDialog(self):
        """Displays a settings dialog window to allow the user to configure various settings.

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
    def set_window_position(window_title, x, y, width, height):
        """Set the position and size of a window by its title.

            Args:
                window_title (str): The title of the window to position.
                x (int): The x-coordinate of the window's new position.
                y (int): The y-coordinate of the window's new position.
                width (int): The new width of the window.
                height (int): The new height of the window.

            Note:
                This method works only on Windows systems using the win32gui library.
        """
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetWindowPos(hwnd, win32gui.HWND_TOP, x, y, width, height, win32gui.SWP_SHOWWINDOW)

    def activate_after_effects_window(self):
        """Activates and positions the After Effects window.

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
                print(f"Error managing After Effects window: {e}")

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
        self.run_after_effects_button.setIcon(colorize_svg("images/after-effects.png", Qt.white))  # Set the color here
        self.run_after_effects_button.setIconSize(QSize(20, 20))  # Set icon size
        self.run_after_effects_button.setFixedSize(80, 80)

        self.open_project_button = QPushButton()
        self.open_project_button.setText("Open Project")
        self.open_project_button.setIcon(colorize_svg("images/book-open.svg", Qt.white))  # Set the color here
        self.open_project_button.setIconSize(QSize(20, 20))  # Set icon size (optional)
        self.open_project_button.setFixedSize(80, 80)

        self.new_project_button = QPushButton()
        self.new_project_button.setText("New Project")
        self.new_project_button.setIcon(colorize_svg("images/file-plus.svg", Qt.white))  # Set the color here
        self.new_project_button.setIconSize(QSize(20, 20))  # Set icon size (optional)
        self.new_project_button.setFixedSize(80, 80)

        self.import_file_button = QPushButton()
        self.import_file_button.setText("Import Asset")
        self.import_file_button.setIcon(colorize_svg("images/plus-square.svg", Qt.white))  # Set the color here
        self.import_file_button.setIconSize(QSize(20, 20))  # Set icon size (optional)
        self.import_file_button.setFixedSize(80, 80)

        self.settings_button = QPushButton("Settings")
        self.settings_button.setIcon(QIcon('images/settings.svg'))  # Ensure you have an appropriate icon
        self.settings_button.setIconSize(QSize(20, 20))
        self.settings_button.setFixedSize(200, 40)

        self.run_after_effects_button.setFixedSize(200, 40)
        self.open_project_button.setFixedSize(200, 40)
        self.new_project_button.setFixedSize(200, 40)
        self.import_file_button.setFixedSize(200, 40)

        # Create a table widget to display last project files
        self.last_projects_table = QTableWidget()
        self.last_projects_table.setColumnCount(5)  # Increase column count for buttons
        self.last_projects_table.setHorizontalHeaderLabels(['Project Name', 'File Path', 'Last Modified', 'Open Project', 'Delete Project'])
        self.last_projects_table.horizontalHeader().setStretchLastSection(True)
        self.last_projects_table.setSortingEnabled(True)
        self.last_projects_table.setSelectionBehavior(QTableWidget.SelectRows)

        # Create a table widget to display imported assets
        self.imported_assets_table = QTableWidget()
        self.imported_assets_table.setColumnCount(5)  # Adjusted from 4 to 5
        self.imported_assets_table.setHorizontalHeaderLabels(['Asset Name', 'File Path', 'Last Modified', 'Imported By', 'Actions'])
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
        projects_layout.addWidget(self.last_projects_table)

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

    def filter_projects(self, text):
        """
        Filters displayed projects based on search input.

        Args:
        text (str): Search query entered by the user.
        """
        if not text:
            self.last_project_files = self.load_all_projects()  # Reset to all projects
            self.populate_last_projects_table()
            return

        filtered_projects = [project for project in self.last_project_files if text.lower() in project.name.lower()]
        self.last_project_files = filtered_projects
        self.populate_last_projects_table()

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

    def load_all_projects(self):
        """
        Loads all previously opened projects from application settings.

        Retrieves and returns a list of Path objects representing previously opened and saved projects.

        Returns:
        List[Path]: A list of Path objects for each previously opened project.
        """
        settings = QSettings("YourOrganization", "AfterEffectsPipeline")
        loaded_last_proj_dir = settings.value("last_project_directory", None)
        if not loaded_last_proj_dir:
            return []
        last_project_files = []
        for i in settings.value("last_project_files", []):
            last_project_files.append(Path(i))
        return last_project_files

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
        self.settings_button.clicked.connect(self.showSettingsDialog)

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Input Connections
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    def open_project_connection(self, project_path: str = None):
        """
        Opens an After Effects project.

        Opens the specified After Effects project or the project selected in the last projects table if no path is provided.
        It handles the subprocess call to After Effects with the project path.

        Args:
            project_path (str, optional): The file path to the After Effects project to be opened. If None, the project
                                          selected in the last projects table will be opened. Defaults to None.
        """
        if project_path:
            aerender_path = r"C:\Program Files\Adobe\Adobe After Effects 2024\Support Files\AfterFX.exe"
            command = [aerender_path, project_path]

            # Execute the After Effects command asynchronously
            self.run_after_effects_command_async(command)
        else:
            selected_row = self.last_projects_table.currentRow()
            if selected_row == -1:
                self.show_error_message("No project selected", "Please select a project to open.")
                return
            project_path_item = self.last_projects_table.item(selected_row, 1)
            if project_path_item:
                project_path = project_path_item.text()
                aerender_path = r"C:\Program Files\Adobe\Adobe After Effects 2024\Support Files\AfterFX.exe"
                command = [aerender_path, project_path]

                # Call the asynchronous method to run After Effects command
                self.run_after_effects_command_async(command)
            else:
                self.show_error_message("Invalid project", "Selected project is invalid.")
                return

    def delete_project(self, index: int):
        """
        Deletes a project file from the filesystem and updates the UI.

        Args:
            index (int): The index of the project file in the last project files list to be deleted.
        """
        if 0 <= index < len(self.last_project_files):
            project_info = self.last_project_files[index]
            project_path_str = project_info["path"]
            project_path = Path(project_path_str)

            confirmation = QMessageBox.question(self, "Confirmation", "Are you sure you want to delete this project?",
                                                QMessageBox.Yes | QMessageBox.No)
            if confirmation == QMessageBox.Yes:
                try:
                    # Check if the project file exists and delete it
                    if project_path.exists():
                        os.remove(project_path)
                        # Now remove this project from the list of projects
                        del self.last_project_files[index]
                        # Update the table to reflect the deletion
                        self.populate_last_projects_table()
                        # Save the updated list of projects to persistent storage, if applicable
                        self.save_settings()
                    else:
                        QMessageBox.information(self, "File Not Found", "The project file does not exist.")
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Error deleting project file: {e}")
        else:
            QMessageBox.warning(self, "Invalid Selection", "The selected project is invalid.")

    def create_new_project_connection(self):
        # Use the user-defined projects folder
        projectsFolder = QSettings("YourOrganization", "AfterEffectsPipeline").value("projectsFolder", "default/path/to/projects")
        if not os.path.exists(projectsFolder):
            os.makedirs(projectsFolder, exist_ok=True)

        try:
            project_name, ok_pressed = QInputDialog.getText(self, "New Project", "Enter the project name:")
            if ok_pressed and project_name:
                new_project_path = Path(projectsFolder, f"{project_name}.aep")
                template_project_path = "C:/Users/marty/ae_test_2/Blank_Project.aep"
                shutil.copyfile(template_project_path, new_project_path)

                # Fetch the current operating system's username
                created_by = getpass.getuser()  # Use the OS username as the project creator's name

                self.update_last_projects_list(new_project_path, created_by)
                self.save_settings()
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
        aerender_path = r"C:\Program Files\Adobe\Adobe After Effects 2024\Support Files\AfterFX.exe"
        if aerender_path:
            command = [aerender_path]
            try:
                # Use subprocess.Popen with start method
                subprocess.Popen(command, start_new_session=True)
            except Exception as e:
                print(f"Error launching After Effects: {e}")
        else:
            self.show_error_message("After Effects path not provided", "Please enter the path to After Effects.")

    def open_after_effects_project_connection(self):
        """
        Opens a selected After Effects project file.

        Prompts the user to select an After Effects project file via a file dialog. Updates the last project directory
        and list based on the selection, saves the updated information to settings, and launches After Effects with the
        selected project.
        """
        aerender_path = r"C:\Program Files\Adobe\Adobe After Effects 2024\Support Files\AfterFX.exe"
        if not Path(aerender_path).exists():
            self.show_error_message("After Effects path not provided", "Executable not found.")
            return

        initial_dir = str(self.last_project_directory) if self.last_project_directory else ""
        project_path, _ = QFileDialog.getOpenFileName(self, "Open After Effects Project", initial_dir,
                                                      "After Effects Project Files (*.aep);;All Files (*)")
        if project_path:
            print(f"Project selected: {project_path}")  # Diagnostic print
            self.last_project_directory = Path(project_path).parent
            self.update_last_projects_list(Path(project_path))
            self.save_settings()

            # Use QProcess to start After Effects with the selected project
            self.process = QProcess(self)
            self.process.setProgram(aerender_path)
            self.process.setArguments([project_path])

            # Optional: Connect signals for process start, finish, and error handling
            self.process.started.connect(lambda: print("After Effects started."))
            self.process.finished.connect(lambda exitCode: print(f"After Effects finished with exit code {exitCode}."))
            self.process.errorOccurred.connect(lambda error: print(f"Error occurred: {self.process.errorString()}"))

            self.process.start()  # Start the process
            print("Command to open After Effects has been issued.")

    def onProcessFinished(self):
        print("After Effects command executed successfully.")

    def is_after_effects_running(self):
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
        """
        Imports a selected file into the current After Effects project, updating the UI and settings accordingly.
        """
        if not self.is_after_effects_running():
            self.show_error_message("After Effects Not Running", "Adobe After Effects must be open to import assets.")
            return

        aePath = self.get_after_effects_path()
        if not aePath:
            return

        assetsFolder = self.get_assets_folder_path()
        if not assetsFolder:
            return

        file_name = self.prompt_user_for_file(assetsFolder)
        if not file_name:
            return

        # Copy the file to the assets folder and prepare for import
        destination = shutil.copy(file_name, assetsFolder)
        imported_by = getpass.getuser()
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Generate and execute JSX script for importing
        self.execute_jsx_script(aePath, destination)

        # Update UI and settings
        self.update_imported_assets(file_name, current_time, imported_by)

    def get_after_effects_path(self):
        aePath = QSettings("YourOrganization", "AfterEffectsPipeline").value("aePath", "")
        if not aePath or not os.path.isfile(aePath):
            self.show_error_message("Error", "After Effects executable path is not set correctly in settings.")
            return None
        return aePath

    def get_assets_folder_path(self):
        assetsFolder = QSettings("YourOrganization", "AfterEffectsPipeline").value("assetsFolder", "")
        if not assetsFolder or not os.path.exists(assetsFolder):
            self.show_error_message("Assets folder not set", "Please set the assets folder path in settings.")
            return None
        return assetsFolder

    def prompt_user_for_file(self, assetsFolder):
        file_name, _ = QFileDialog.getOpenFileName(self, "Import File into After Effects", assetsFolder, "All Files (*)")
        return file_name if file_name else None

    def execute_jsx_script(self, aePath, file_path):
        corrected_path = file_path.replace('\\', '\\\\')
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
        self.run_after_effects_command_async([aePath, '-r', script_file_path])

    def update_imported_assets(self, file_path, last_modified, imported_by):
        asset_name = os.path.basename(file_path)
        self.imported_assets.append({"path": file_path, "last_modified": last_modified, "imported_by": imported_by})
        self.populate_imported_assets_table(asset_name, file_path, last_modified, imported_by)
        self.save_imported_assets_settings()

    def populate_imported_assets_table(self, asset_name, file_path, last_modified, imported_by):
        """
        Populates the imported assets table with the new asset.

        Args:
            asset_name (str): The name of the imported asset.
            file_path (str): The file path of the imported asset.
            last_modified (str): Last modified timestamp of the asset.
        """
        row_count = self.imported_assets_table.rowCount()
        self.imported_assets_table.insertRow(row_count)
        self.imported_assets_table.setItem(row_count, 0, QTableWidgetItem(asset_name))
        self.imported_assets_table.setItem(row_count, 1, QTableWidgetItem(file_path))
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

    def save_imported_assets_settings(self):
        """
        Save the list of imported assets to settings.

        This method saves the list of imported assets to the application settings.
        It uses QSettings to store the data persistently.
        """
        settings = QSettings("YourOrganization", "AfterEffectsPipeline")
        # Serialize imported_assets list into a JSON string
        serialized_assets = json.dumps(self.imported_assets)
        settings.setValue("imported_assets", serialized_assets)

    def import_to_current_project(self, file_path):
        """
        Imports the selected file into the current After Effects project.

        Args:
            file_path (str): The file path of the asset to import into After Effects.
        """

        if not self.is_after_effects_running():
            self.show_error_message("After Effects Not Running", "Adobe After Effects must be open to perform this action.")
            return

        aerender_path = r"C:\Program Files\Adobe\Adobe After Effects 2024\Support Files\AfterFX.exe"
        if not aerender_path:
            self.show_error_message("After Effects path not provided", "Please enter the path to After Effects.")
            return

        script_content = f"""
        var proj = app.project;
        if (proj) {{
            var fileToImport = new ImportOptions();
            fileToImport.file = new File("{file_path}");
            proj.importFile(fileToImport);
        }}
        """
        # Create a temporary JSX script file
        script_file_path = tempfile.mktemp(suffix=".jsx")
        with open(script_file_path, 'w') as script_file:
            script_file.write(script_content)

        # Execute the After Effects command synchronously
        command = [aerender_path, '-r', script_file_path]
        try:
            subprocess.run(command, check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error during After Effects execution: {e}")
        os.remove(script_file_path)

    def delete_imported_asset(self, row):
        confirmation = QMessageBox.question(self, "Confirmation", "Are you sure you want to delete this asset?",
                                            QMessageBox.Yes | QMessageBox.No)
        if confirmation == QMessageBox.Yes:
            asset_name = self.imported_assets_table.item(row, 0).text()
            assetsFolder = QSettings("YourOrganization", "AfterEffectsPipeline").value("assetsFolder", "")
            asset_path = os.path.join(assetsFolder, asset_name)
            self.imported_assets_table.removeRow(row)
            self.imported_assets.pop(row)

            if os.path.exists(asset_path):
                try:
                    os.remove(asset_path)
                    print(f"Deleted asset at: {asset_path}")
                except OSError as e:
                    QMessageBox.critical(self, "Error", f"Failed to delete asset file: {e}")
                    print(f"Failed to delete asset at: {asset_path}. Error: {e}")
            else:
                QMessageBox.information(self, "Error", "Asset file not found in the assets folder.")

            self.save_imported_assets_settings()

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Back-end Functions
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    def delete_project(self, index: int):
        """
        Deletes a project file from the filesystem and updates the UI.

        Args:
            index (int): The index of the project file in the last project files list to be deleted.
        """
        # Validate index range
        if not (0 <= index < len(self.last_project_files)):
            QMessageBox.warning(self, "Invalid Selection", "The selected project is invalid.")
            return

        project_info = self.last_project_files[index]
        project_path = Path(project_info["path"])

        # Confirm deletion with the user
        confirmation = QMessageBox.question(
            self, "Confirmation", "Are you sure you want to delete this project?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if confirmation != QMessageBox.Yes:
            return

        # Attempt to delete the project file
        try:
            if not project_path.exists():
                raise FileNotFoundError("The project file does not exist.")

            os.remove(project_path)
            # Reflect changes in the UI and settings
            del self.last_project_files[index]
            self.populate_last_projects_table()
            self.save_settings()
            QMessageBox.information(self, "Success", "The project has been successfully deleted.")

        except FileNotFoundError as e:
            QMessageBox.information(self, "File Not Found", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error deleting project file: {e}")

    def remove_row(self, index: int):
        """
        Remove a row from the last projects table.

        This method removes a row at the specified index from the last projects table.
        It also updates the project list and saves the changes to the application settings.

        Args:
            index (int): The index of the row to be removed.
        """
        if 0 <= index < len(self.last_project_files):
            del self.last_project_files[index]
            self.populate_last_projects_table()
            self.save_settings()

    def run_after_effects_command_async(self, command: list[str]):
        """
        Runs an After Effects command asynchronously using a QThread.

        Args:
            command (list[str]): The After Effects command to execute asynchronously.
        """
        self.thread = AfterEffectsThread(command)
        self.thread.finished.connect(self.thread_finished)
        self.thread.start()

    def update_last_projects_list(self, project_path: Path, created_by: str):
        """
        Adds a project to the list of last opened projects.

        Args:
            created_by:
            project_path (Path): The file path of the project to add to the list.
        """
        project_info = {"path": project_path, "created_by": created_by}
        self.last_project_files.append(project_info)
        self.populate_last_projects_table()

    def populate_last_projects_table(self):
        """
        Populate the table with the last opened projects, combining "Open Project" and "Delete Project"
        buttons under a single "Actions" category.
        """
        self.last_projects_table.setRowCount(0)  # Reset the table contents

        # Filter out projects that no longer exist on the filesystem
        existing_projects = [project for project in self.last_project_files if Path(project["path"]).exists()]

        # Update the last project files list to only include existing projects
        self.last_project_files = existing_projects

        for index, project_info in enumerate(self.last_project_files):
            project_path = Path(project_info["path"])
            project_name = project_path.name
            file_path = str(project_path)
            last_modified = datetime.fromtimestamp(project_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            created_by = project_info["created_by"]  # Assuming each project_info dict has a "created_by" key

            self.last_projects_table.insertRow(index)
            self.last_projects_table.setItem(index, 0, QTableWidgetItem(project_name))
            self.last_projects_table.setItem(index, 1, QTableWidgetItem(file_path))
            self.last_projects_table.setItem(index, 2, QTableWidgetItem(last_modified))
            self.last_projects_table.setItem(index, 3, QTableWidgetItem(created_by))

            # Container widget for action buttons
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(0, 0, 0, 0)
            open_button = QPushButton('Open')
            delete_button = QPushButton('Delete')

            # Connect the buttons to their respective slot functions
            open_button.clicked.connect(lambda _, idx=index: self.open_project_connection(self.last_project_files[idx]["path"].as_posix()))
            delete_button.clicked.connect(lambda _, idx=index: self.delete_project(idx))

            actions_layout.addWidget(open_button)
            actions_layout.addWidget(delete_button)

            self.last_projects_table.setCellWidget(index, 4, actions_widget)

        # Optionally, save the updated list of projects to persist the changes
        self.save_settings()

    @staticmethod
    def thread_finished():
        print("After Effects command executed successfully.")

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

    def load_settings(self):
        """
        Load application settings, including paths and the list of last projects and imported assets,
        with enhanced JSON parsing safety.
        """
        settings = QSettings("YourOrganization", "AfterEffectsPipeline")
        self.last_project_directory = Path(settings.value("last_project_directory", "")) \
            if settings.value("last_project_directory", None) else None

        # Simplify loading and parsing of the last project files
        try:
            self.last_project_files = [
                {"path": Path(project["path"]), "created_by": project.get("created_by", "Unknown")}
                for project in json.loads(settings.value("last_project_files", "[]"))
                if isinstance(project, dict)]
        except json.JSONDecodeError:
            self.last_project_files = []

        # Simplify loading and parsing of the imported assets
        try:
            self.imported_assets = json.loads(settings.value("imported_assets", "[]"))
        except json.JSONDecodeError:
            self.imported_assets = []

        # Clear the tables before repopulating
        self.last_projects_table.setRowCount(0)
        self.imported_assets_table.setRowCount(0)
        self.populate_last_projects_table()
        for asset in self.imported_assets:
            self.populate_imported_assets_table(
                asset_name=os.path.basename(asset["path"]),
                file_path=asset["path"],
                last_modified=asset["last_modified"],
                imported_by=asset.get("imported_by", "Unknown")
            )

    def save_settings(self):
        """
        Saves the application settings, including paths and the list of imported assets.
        """
        settings = QSettings("YourOrganization", "AfterEffectsPipeline")
        # Serialize each project in last_project_files as a dictionary including the path and created_by info
        serialized_projects = json.dumps([{"path": str(project["path"]), "created_by": project["created_by"]} for project in self.last_project_files])
        settings.setValue("last_project_files", serialized_projects)
        # Save the list of imported assets as a JSON string
        settings.setValue("imported_assets", json.dumps(self.imported_assets))


class AfterEffectsThread(QThread):
    finished = pyqtSignal()

    def __init__(self, command: list[str]):
        super().__init__()
        self.command = command

    def run(self):
        try:
            subprocess.run(self.command, check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error during After Effects execution: {e}")
        finally:
            self.finished.emit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AfterEffectsPipeline()
    window.show()

    sys.exit(app.exec_())
