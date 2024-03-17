from cx_Freeze import setup, Executable

# Specify the base directory and script to convert into an executable
base = None
script = 'C:/Users/marty/ae_test_2/pyae_v1.py'  # Adjust the path as per your actual location

# Set up the executable
executables = [Executable(script, base=base, icon='C:/Users/marty/ae_test_2/PyAE_icon.ico')]  # Replace 'PyAE_icon.ico' with your actual icon file name and path

# Additional options
options = {
    'build_exe': {
        'includes': [],
        'include_files': ['C:/Users/marty/ae_test_2/PyAE_icon.ico'],  # Add any additional files or directories here
    }
}

setup(name='py_ae',
      version='1.0',
      description='Description of your app',
      executables=executables,
      options=options)
