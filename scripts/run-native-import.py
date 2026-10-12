"""Limit the Word UI import and leave AppleScript progress visible in Terminal."""
from pathlib import Path
import subprocess
import sys


def run_import(base, module):
    script = Path(__file__).with_name('import-native-vba.applescript')
    try:
        result = subprocess.run(
            ['osascript', str(script), base, module],
            text=True, stdout=subprocess.PIPE, check=True, timeout=60,
        )
    except subprocess.TimeoutExpired:
        raise ValueError('Automation stopped after 60 seconds. Check the last step and the Word dialogs; nothing was installed.') from None
    except subprocess.CalledProcessError:
        raise ValueError('Word interrupted the import; check the message above. Nothing was installed.') from None
    if result.stdout.strip() != 'BIP_VBA_IMPORTED':
        raise ValueError('Word did not confirm the import; nothing was installed.')
    return 'BIP_VBA_IMPORTED'


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise SystemExit('Pass the base template and the VBA module.')
    try:
        print(run_import(sys.argv[1], sys.argv[2]))
    except ValueError as error:
        raise SystemExit(str(error))
