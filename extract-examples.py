import os
import shutil
from pathlib import Path
from zebra_vba_packager import decompile_xl, backup_last_50_paths
from tempfile import TemporaryDirectory, gettempdir
from locate import this_dir

backup_dir = Path(gettempdir(), "BackupVBALib")
with TemporaryDirectory() as tempdir:
    if (ex_path := this_dir().joinpath("add_examples")).exists():
        backup_last_50_paths(backup_dir, ex_path)

    if (xl_path := this_dir().joinpath("VBALib.xlsb")).exists():
        backup_last_50_paths(backup_dir, xl_path)

        decompile_xl(xl_path, tempdir)
        shutil.rmtree(ex_path, ignore_errors=True)
        os.makedirs(ex_path, exist_ok=True)
        for i in Path(tempdir).rglob("examples_*.bas"):
            i.rename(ex_path.joinpath(i.name))