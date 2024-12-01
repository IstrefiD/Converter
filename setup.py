from cx_Freeze import setup, Executable

# Collect required dependencies
build_exe_options = {
    "packages": ["os", "tkinter", "pyexcel_ods", "openpyxl", "xml.etree.ElementTree"],
    "include_files": [],  # Add any additional files if needed
}

# Define the executable
executables = [
    Executable("app.py", base="Win32GUI")  # No targetName argument, just use base="Win32GUI"
]

# Setup function to bundle everything into an exe
setup(
    name="File Converter",
    version="1.0",
    description="A simple file conversion tool",
    options={"build_exe": build_exe_options},
    executables=executables
)
