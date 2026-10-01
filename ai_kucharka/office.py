import hashlib
import logging
import os
import platform
import shutil
import subprocess

import psutil

if platform.system() == "Windows":
    import win32com.client as win32


class MSOffice:
    def __init__(self, logger: logging.Logger) -> None:
        os.environ["EXCEL_DISABLE_APPADDINS"] = "1"

        self.logger = logger

        self.word = win32.Dispatch("Word.Application")
        self.word.Visible = False
        self.word.DisplayAlerts = False
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False

        # Disconnect Excel macros
        self.excel.AutomationSecurity = 3
        try:
            for addin in self.excel.COMAddIns:
                addin.Connect = False
        except Exception:
            pass

    def close(self) -> None:
        self.word.Quit()
        self.excel.Quit()

    def convert_file(self, input_path: str, output_path: str) -> None:
        output_path = os.path.abspath(output_path)

        _, ext = os.path.splitext(input_path.lower())
        if ext not in [".doc", ".xls"]:
            raise ValueError("Unsupported extension. Should be doc or xls only.")

        doc = None
        try:
            if ext == ".doc":
                doc = self.word.Documents.Open(input_path, ReadOnly=True)
                doc.SaveAs(output_path, FileFormat=16)
            elif ext == ".xls":
                doc = self.excel.Workbooks.Open(input_path, ReadOnly=True, CorruptLoad=1)
                doc.SaveAs(output_path, FileFormat=51)
        except Exception as e:
            self.logger.warning("Conversion failed", extra={"file": input_path, "type": "failed_conversion", "error": str(e)})
        finally:
            if doc is not None:
                try:
                    doc.Close()
                except Exception:
                    pass

    @classmethod
    def kill_all(cls) -> None:
        for proc in psutil.process_iter(["name"]):
            if proc.info["name"] in ("EXCEL.EXE", "WINWORD.EXE"):
                try:
                    proc.kill()
                except psutil.NoSuchProcess:
                    pass


def convert_office_file(
    input_path: str,
    output_dir: str | None = None,
    output_file_name: str | None = None,
    office: MSOffice | None = None,
) -> str:

    basename = os.path.basename(input_path)
    name, ext = os.path.splitext(basename)
    ext = ext.lower()

    if ext not in [".doc", ".xls"]:
        raise ValueError("Only .doc and .xls files can be converted")

    # Output folder
    if output_dir is None:
        output_dir = os.path.dirname(input_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # Output filename
    if output_file_name is None:
        hash_id = hashlib.sha256(input_path.encode()).hexdigest()[:16]
        output_file_name = f"{hash_id}{ext}x"
    output_path = os.path.join(output_dir, output_file_name)

    # Skip if already exists and was not modified later
    if os.path.exists(output_path) and os.path.getmtime(output_path) > os.path.getmtime(input_path):
        return output_path

    # Convert the file
    system = platform.system()
    if system in ["Linux", "Darwin"]:
        subprocess.call(["soffice", "--headless", "--convert-to", ext + "x", "--outdir", output_dir, input_path])
        expected_output = os.path.join(output_dir, f"{name}{ext}x")
        if os.path.exists(expected_output):
            shutil.move(expected_output, output_path)
        else:
            raise FileNotFoundError(f"LibreOffice did not produce: {expected_output}")
    elif system == "Windows":
        if office is None:
            raise ValueError("office instance required on Windows")
        office.convert_file(input_path, output_path)
    else:
        raise OSError("Unsupported OS")

    return output_path
