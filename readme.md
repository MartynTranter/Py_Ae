# PyAE - After Effects Pipeline

## Overview

PyAE - After Effects Pipeline is a robust desktop application designed to optimize the workflow of Adobe After Effects users. Utilizing the powerful PyQt5 framework, this application introduces an intuitive approach to project and asset management. Its aim is to streamline the creative process for video editors and motion graphics artists, ensuring an organized, efficient workflow.

## Features

PyAE comes packed with features aimed at simplifying the After Effects workflow:

- [x] **Open After Effects Projects**: Directly open projects within After Effects with a simple click.
- [x] **Create New Projects**: Quickly start new projects using customizable templates.
- [x] **Import Files**: Import various media files directly into your active After Effects project.
- [x] **Run After Effects**: Launch After Effects directly from the application for a seamless workflow.
- [ ] **Version Control**: *Upcoming* - Manage different versions of your After Effects projects efficiently.
- [ ] **Batch Importer**: *Planned* - Import multiple files simultaneously with batch processing capabilities.
- [ ] **Send Project to Premiere**: *In Development* - Easily transfer projects between After Effects and Premiere Pro for an integrated workflow.
- [ ] **Light and Dark Mode**: *Future Feature* - Customize the application's appearance with light and dark themes.

## Getting Started

### Prerequisites

To get started with PyAE - After Effects Pipeline, ensure you have the following installed:

- Python 3.6 or later for the core application.
- PyQt5 for the graphical user interface.
- psutil for managing system processes.
- pygetwindow and pywin32 for enhanced window management on Windows platforms.

### Installation

Follow these steps to install PyAE - After Effects Pipeline:

1. **Clone the repository**:
    ```bash
    git clone https://yourrepository/AfterEffectsPipeline.git
    ```
    
2. **Navigate to the project directory**:
    ```bash
    cd AfterEffectsPipeline
    ```
    
3. **Install the required dependencies**:
    ```bash
    pip install PyQt5 psutil pygetwindow pywin32
    ```

### Running the Application

Launch PyAE - After Effects Pipeline using the command:
```bash
python after_effects_pipeline.py
```

## Usage

### Configuration

Upon first launch, navigate to the settings via the "Settings" button to configure the necessary paths for:

- After Effects executable
- Assets folder
- Projects folder

These paths are essential for the application to function correctly with your After Effects environment.

### Managing Projects

- **Create a New Project**: Click on "New Project", enter a name, and a project template will be used to create a new After Effects project.
- **Open an Existing Project**: Select a project from the list and click "Open Project" to launch it in After Effects.
- **Delete a Project**: Select a project and choose "Delete Project" to remove it from your filesystem.

### Importing Assets

With After Effects running, you can import assets directly into your active project:

1. Click "Import Asset".
2. Navigate to the asset you wish to import.
3. Select the asset and confirm to import it into the currently open After Effects project.

## Customization

To customize the application's appearance, modify the `Combinear.qss` file located in the project's root directory. This file contains the stylesheet rules used by the PyQt5 application.

## Contributing

We welcome contributions to After Effects Pipeline. To contribute:

1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/AmazingFeature`).
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`).
4. Push to the branch (`git push origin feature/AmazingFeature`).
5. Open a pull request.

## License

This project is licensed under the MIT License - see the `LICENSE.md` file for details.

## Contact

- Twitter - [@martyntranter](https://twitter.com/martyntranter)
- Project Link: [https://github.com/yourrepository/AfterEffectsPipeline](https://github.com/yourrepository/AfterEffectsPipeline)
