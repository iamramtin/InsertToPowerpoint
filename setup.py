from cx_Freeze import setup, Executable

base = None

executables = [Executable("run.py", base=base)]

packages = ["idna", "os", "tkinter", "docx", "pptx"]
options = {
    "build_exe": {
        "packages": packages,
    },
}

setup(
    name="Blacfox-Client",
    options={"build.exe": {"packages": ["tkinter", "docx", "pptx"], "include_files": ["Blacfox Template.pptx"]}},
    version="0.1",
    description="Converting *.docx to *.pptx.",
    executables=executables
)
