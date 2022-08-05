from copy import copy
from tempfile import gettempdir, TemporaryDirectory

from zebra_vba_packager import (
    Config,
    Source,
    compile_xl,
    runmacro_xl,
    pack,
    backup_last_50_paths,
)
from locate import this_dir
from textwrap import dedent
import os
import shutil
from contextlib import suppress
from pathlib import Path


backup_dir = Path(gettempdir(), "BackupVBALib")
with TemporaryDirectory() as tempdir:
    if (f := this_dir().joinpath("VBALib")).exists():
        backup_last_50_paths(backup_dir, f)
    if (f := this_dir().joinpath("VBALib.xlsb")).exists():
        backup_last_50_paths(backup_dir, f)


output = this_dir().joinpath("VBALib")


    if isinstance(s, bytes):
        return s.replace(b"\r\n", b"\n")
    else:
        return s.replace("\r\n", "\n")


def mid_process(source):
    for pth in source.temp_transformed.rglob("*.bas"):
        with pth.open("rb") as f:
            txt_lines = to_ln(f.read()).split(b"\n")

        do_overwrite = False
        for i, line in enumerate(txt_lines):
            if line.strip().startswith(b"Public"):
                if line.strip().split()[2].lower() == b"as":
                    do_overwrite = True
                    ii = line.lower().find(b"public")
                    line = line[0:ii] + b"Private" + line[ii + len(b"public") :]
                    txt_lines[i] = line

        if do_overwrite:
            with pth.open("wb") as f:
                f.write((b"\r\n".join(txt_lines)))


def office64_ptr_compatability(bin_lines):
    bin_lines = copy(bin_lines)

    # strip trailing ending lines
    i = -1
    while (i := i + 1) < len(bin_lines):
        if bin_lines[i].strip().lower().split()[1:2] == [b"declare"]:
            j = i
            for j in range(i, len(bin_lines)):
                if not bin_lines[j].strip().endswith(b"_"):
                    break
            line = b" ".join(
                [bin_lines[ii].strip().rstrip(b"_") for ii in range(i, j + 1)]
            )

            bin_lines[i] = line
            for ii in range(i, j):
                bin_lines.pop(i + 1)

    i = -1
    while (i := i + 1) < len(bin_lines):
        if b'declare' in bin_lines[i].lower():
            print(bin_lines[i])
        if bin_lines[i].strip().lower().split()[1:2] == [b"declare"]:
            txt64 = (
                bin_lines[i]
                .replace(b" Declare ", b" Declare PtrSafe ")
                .replace(b"Long", b"LongPtr")
            )
            all_pointer_declare.append(
                dedent(
                    f"""
            #If VBA7 Then
                {txt64.decode('utf-8')}
            #Else
                {bin_lines[i].decode('utf-8')}
            #End If
            """
                )
                .strip()
                .encode("utf-8")
            )

            bin_lines.pop(i)
            i -= 1

    return bin_lines


def functions_start(bin_lines):
    i = -1
    while (i := i + 1) < len(bin_lines):
        funcdecl = bin_lines[i].lower().strip().split()[:2]
        if (
            funcdecl[:1] == [b"function"]
            or funcdecl[:1] == [b"sub"]
            or funcdecl == [b"private", b"function"]
            or funcdecl == [b"private", b"sub"]
            or funcdecl == [b"public", b"function"]
            or funcdecl == [b"public", b"sub"]
        ):
            break

    return i


# Seperate step for this library
common_lib = Config(
    Source(
        git_source="https://github.com/nylen/vba-common-library.git",
        git_rev="1e21b0d",  # 2014
        glob_include=["*.cls", "*.bas"],
        glob_exclude=["*VBALib_VERSION*"],
        auto_bas_namespace=False,
        auto_cls_rename=False,
    )
)
common_lib.run()

all_bas_lines = []
all_pointer_declare = []
all_precode_declare = []
for fpath in common_lib.output_dir.glob("*.bas"):
    if "z__" in str(fpath) and str(fpath).lower().endswith(".bas"):
        continue

    with fpath.open("rb") as f:
        txt_lines = to_ln(f.read()).split(b"\n")

    txt_lines = office64_ptr_compatability(txt_lines)
    i = functions_start(txt_lines)

    all_precode_declare.extend(
        [j for j in txt_lines[1:i] if not j.lower().strip().startswith(b"option ")]
    )
    all_bas_lines.extend(txt_lines[i:])
    os.remove(fpath)

for fpath in common_lib.output_dir.glob("*.cls"):
    with fpath.open("rb") as f:
        txt = f.read()

    with fpath.open("wb") as f:
        f.write(
            txt.replace(b"Attribute VB_Exposed = False", b"Attribute VB_Exposed = True")
        )

frm1 = b"    Dim bytesRead As Long"
to1 = b"""
    #If VBA7 Then
        Dim bytesRead As LongPtr
    #Else
        Dim bytesRead As Long
    #End If
"""

frm2 = b"    Dim ret As Long"
to2 = b"""
    #If VBA7 Then
        Dim ret As LongPtr
    #Else
        Dim ret As Long
    #End If
"""

with common_lib.output_dir.joinpath("concatenated.bas").open("wb") as f:
    f.write(
        b"\r\n".join(
            [b'Attribute VB_Name = "VLib"']
            + [b"Option Explicit"]
            + all_pointer_declare
            + all_precode_declare
            + all_bas_lines
        )
        .replace(frm1, to1)
        .replace(frm2, to2)
    )

# Find all function/sub names and attach "VLib." in front ot them
vlib_funcnames = set()
for i in all_bas_lines:
    fline = i.strip().split()
    with suppress(IndexError):
        if fline[0].lower() in (b"function", b"sub") or fline[1].lower() in (
            b"function",
            b"sub",
        ):
            vlib_funcnames.add(fline[1].decode("utf-8").split("(")[0])
            vlib_funcnames.add(fline[2].decode("utf-8").split("(")[0])

vlib_namespace_fix = [(i, "VLib." + i) for i in vlib_funcnames]
vlib_renames = [
    (
        lambda x: x.lower().startswith("vbalib_excel"),
        lambda x: "z_VLib" + x[len("vbalib_excel") :],
    ),
    (
        lambda x: x.lower().startswith("vbalib_"),
        lambda x: "z_VLib" + x[len("vbalib_") :],
    ),
]

common_lib_output_dir2 = Path(str(common_lib.output_dir) + "2")
shutil.rmtree(common_lib_output_dir2, ignore_errors=True)
shutil.copytree(common_lib.output_dir, common_lib_output_dir2)


# -----------------------------------------------------------------------
# Aggregate all the rest of the sources
# -----------------------------------------------------------------------
Config(
    ## TODO: why doesn't auto_cls_rename work?
    # Source(
    #    git_source="https://github.com/sancarn/stdVBA.git",
    #    git_rev="50fdb99",  # 2021-07
    #    glob_include=['src/*.cls'],
    #    auto_cls_rename=True
    # ),
    Source(
        git_source="https://github.com/VirtualActuary/SHA256-and-MD5-for-VBA.git",
        glob_include=["*.cls"],
        rename_overwrites={"HS256": "z__Hash"},
    ),
    Source(
        git_source="https://github.com/ws-garcia/VBA-CSV-interface.git",
        git_rev="1b23247",  # 2022-03
        glob_include=["**/src/*.cls"],
        rename_overwrites={
            "CSVArrayList": "z_WsCsvArrayList",
            "CSVcallBack": "z_WsCsvCallBack",
            "CSVdialect": "z_WsCsvDialect",
            "CSVexpressions": "z_WsCsvExpressions",
            "CSVparserConfig": "z_WsCsvParserConfig",
            "CSVSniffer": "z_WsCsvSniffer",
            "CSVTextStream": "z_WsCsvTextStream",
            "CSVudFunctions": "z_WsCsvUdFunctions",
            "CSVinterface": "z_WsCsvInterface",  # useful
        },
    ),
    Source(
        git_source="https://github.com/GustavBrock/VBA.Compress.git",
        git_rev="052b889",  # 2020
        glob_include=["**/*.bas"],
        mid_process=mid_process,
        rename_overwrites={
            "FileCompress": "Compress",
        },
    ),
    # The following two projects are dependant on each other:
    Source(
        git_source="https://github.com/VBA-tools/VBA-JSON.git",
        git_rev="v2.3.1",  # 2019
        glob_include=["**/JsonConverter.bas"],
        mid_process=mid_process,
        rename_overwrites={
            "JsonConverter": "Json",  # bas file
        },
    ),
    Source(
        git_source="https://github.com/sdkn104/VBA-CSV.git",
        git_rev="48d98d6",  # 2020
        glob_include=["**/CSVUtils.bas"],
        mid_process=mid_process,
        rename_overwrites={
            "CSVUtils": "CsvUtils",
        },
    ),
    Source(
        git_source="https://github.com/todar/VBA-Strings",
        git_rev="6d25dad",  # 2020
        glob_include=["*.bas"],
        rename_overwrites={"StringFunctions": "StrUtils"},
    ),
    Source(
        path_source=common_lib.output_dir,
        glob_include=["*.cls"],
        glob_exclude=["z__*"],
        rename_overwrites=vlib_namespace_fix + vlib_renames,
    ),
    Source(
        path_source=common_lib_output_dir2,
        glob_include=["*.bas"],
        glob_exclude=["z__*"],
        rename_overwrites=vlib_renames,
    ),
    Source(
        git_source="https://github.com/VirtualActuary/MiscVBAFunctions.git",
        git_rev="8e5e8f3",
        glob_include=[
            "MiscVBAFunctions/**/*.bas",
            "MiscVBAFunctions/**/*.cls",
            "**/thisworkbook.txt",
        ],
        glob_exclude=["**/Test__*"],
        combine_bas_files="Fn",
        auto_bas_namespace=True,
        auto_cls_rename=False,
    ),
    #Source(path_source=this_dir().joinpath("add_examples"), auto_bas_namespace=False),
).run(output)

# Add empty xlsx file
shutil.copy2(
    empty_src := this_dir().joinpath("add_empty_workbook", "empty.xlsx"),
    empty_dest := output.joinpath("empty.xlsx"),
)

# Add examples
shutil.copytree(
    this_dir().joinpath("add_examples"),
    examples_dest := output.joinpath("examples")
)

# Make library into an .xlsb file
compile_xl(output, xl_final := this_dir().joinpath("VBALib.xlsb"))
runmacro_xl(xl_final)

os.remove(empty_dest)
shutil.rmtree(examples_dest)

with suppress(OSError):
    os.remove(zip_dest := this_dir().joinpath("VBALib.zip"))

pack(output, zip_dest, compression_type="zip")


"""
Usefull VBA projects to choose from
https://github.com/sancarn/stdVBA.git (looks promising)
https://github.com/sdkn104/VBA-CSV
https://github.com/VBA-tools
https://github.com/GustavBrock/VBA.Compress
https://github.com/AllenMattson/VBA (???)
https://github.com/carvetighter/VBA-Code-Library
https://github.com/Zadigo/vba_codes
https://github.com/topics/vba-modules (further collection)
https://github.com/Greedquest/VBA-Toolbox (Next level stuff, but bit risky: 
https://github.com/Greedquest/VBA-Toolbox/blob/master/ToolboxSource/TextWriter.cls
https://github.com/nylen/vba-common-library (maybe use tables?)
https://github.com/vbaidiot/ariawase
https://github.com/omegastripes/VBA-JSON-parser
https://github.com/VBA-tools/VBA-JSON.git
https://github.com/x-vba/xlib <--------------- << Combine into single namespace >>
"""
