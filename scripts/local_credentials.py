"""Read and write private credentials for the accounting workflow."""
import json
from pathlib import Path
from tempfile import NamedTemporaryFile


from project_paths import PROJECT_ROOT


CONFIG = PROJECT_ROOT / 'config'


def read_credential(name):
    path = CONFIG / name
    try:
        with path.open(encoding='utf-8') as file:
            return json.load(file)
    except FileNotFoundError:
        return None
    except ValueError as error:
        raise ValueError(f'The credential file {name} is unreadable.') from error


def write_credential(name, value):
    CONFIG.mkdir(mode=0o700, exist_ok=True)
    CONFIG.chmod(0o700)
    staged = None
    try:
        with NamedTemporaryFile(mode='w', encoding='utf-8', dir=CONFIG, delete=False) as file:
            staged = Path(file.name)
            json.dump(value, file)
            file.write('\n')
        staged.replace(CONFIG / name)
    finally:
        if staged is not None:
            staged.unlink(missing_ok=True)
