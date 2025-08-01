import subprocess
import os
import tempfile
import shutil
import argparse
from typing import Optional, Union

# Optional Pillow import (loaded only if needed)
PILLOW_AVAILABLE = False
try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    pass

# Detect PowerShell versions
def detect_powershell_versions():
    ps51_path = shutil.which("powershell.exe")
    ps7_path = shutil.which("pwsh.exe")
    return ps51_path, ps7_path

PS51_PATH, PS7_PATH = detect_powershell_versions()

# Load PowerShell scripts from separate files
def load_ps_script(filename):
    script_path = os.path.join(os.path.dirname(__file__), filename)
    with open(script_path, 'r', encoding='utf-8') as f:
        return f.read()

PS_SCRIPT_51 = load_ps_script('ps_script_51.ps1')
PS_SCRIPT_7 = load_ps_script('ps_script_7.ps1')

class WinOCRSession:
    """Manages a persistent PowerShell session for OCR."""
    def __init__(self, ps_version: str = 'auto', language: str = 'en-US'):
        self.language = language
        self.ps_exe = self._select_ps_exe(ps_version)
        self.ps_script = PS_SCRIPT_7 if 'pwsh' in self.ps_exe else PS_SCRIPT_51
        self.process = None
        self._initialize_session()

    def _select_ps_exe(self, ps_version):
        if ps_version == 'auto':
            if PS7_PATH:
                return PS7_PATH
            elif PS51_PATH:
                return PS51_PATH
            else:
                raise RuntimeError("No PowerShell found.")
        elif ps_version == '7' and PS7_PATH:
            return PS7_PATH
        elif ps_version == '5.1' and PS51_PATH:
            return PS51_PATH
        else:
            raise ValueError(f"PowerShell {ps_version} not available.")

    def _initialize_session(self):
        self.process = subprocess.Popen(
            [self.ps_exe, '-NoExit', '-NoProfile', '-Command', '-'],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8'
        )
        init_code = self.ps_script.split('function Convert-PsoImageToText')[0]
        self.process.stdin.write(init_code + '\n')
        self.process.stdin.flush()
        # Clear startup output
        while self.process.stdout.readline().strip():
            pass

    def ocr_image(self, image: Union[str, 'Image.Image']) -> Optional[str]:
        """Performs OCR on an image path or Pillow Image."""
        if isinstance(image, str):
            image_path = os.path.abspath(image)
        else:
            if not PILLOW_AVAILABLE:
                raise ImportError("Pillow is required for Image objects. Install with 'pip install pillow'.")
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                image.save(temp_file.name)
                image_path = temp_file.name

        try:
            ocr_command = f"Convert-PsoImageToText -Path '{image_path}' -Language '{self.language}'"
            self.process.stdin.write(ocr_command + '\n')
            self.process.stdin.flush()

            output = ''
            while True:
                line = self.process.stdout.readline().strip()
                if not line:
                    break
                output += line + '\n'

            if 'Pillow.Image.Image' in str(type(image)):
                os.unlink(image_path)

            return output.strip()
        except Exception as e:
            print(f"OCR Error: {e}")
            return None

    def close(self):
        """Closes the PowerShell session."""
        if self.process:
            self.process.stdin.write('exit\n')
            self.process.stdin.flush()
            self.process.terminate()
            self.process = None

def ocr_cli():
    """Command-line entry point."""
    parser = argparse.ArgumentParser(description="Windows OCR tool using PowerShell.")
    parser.add_argument('image_paths', nargs='+', help="Image file paths.")
    parser.add_argument('--ps_version', default='auto', choices=['auto', '5.1', '7'], help="PowerShell version.")
    parser.add_argument('--language', default='en-US', help="OCR language.")
    args = parser.parse_args()

    session = WinOCRSession(ps_version=args.ps_version, language=args.language)
    try:
        for path in args.image_paths:
            text = session.ocr_image(path)
            if text:
                print(f"Extracted Text from {path}:\n{text}\n")
            else:
                print(f"Failed to extract text from {path}.")
    finally:
        session.close()

if __name__ == "__main__":
    ocr_cli()
