from zebra_vba_packager import Config, Source
from locate import this_dir

output = this_dir().joinpath("output")

def mid_process(source):
    for pth in source.temp_transformed.rglob("*.bas"):
        with pth.open("rb") as f:
            txt_lines = f.read().split(b"\r\n")

        do_overwrite = False
        for i, line in enumerate(txt_lines):
            if line.strip().startswith(b"Public"):
                if line.strip().split()[2].lower() == b"as":
                    do_overwrite = True
                    ii = line.lower().find(b"public")
                    line = line[0:ii] + b"Private" + line[ii+len(b"public"):]
                    txt_lines[i] = line

        if do_overwrite:
            with pth.open("wb") as f:
                f.write((b"\r\n".join(txt_lines)))


Config(
    Source(
        git_source="https://github.com/ws-garcia/VBA-CSV-interface.git",
        git_rev="v3.1.0",
        glob_include=['**/src/*.cls'],
        rename_overwrites={
            "ECPArrayList": "zWsArray",
            "ECPTextStream": "zWsStream",
            "parserConfig": "zWsCsvConf",
            "CSVinterface": "z__WsCsv__",  # useful
        }
    ),
    Source(
        git_source="https://github.com/GustavBrock/VBA.Compress.git",
        git_rev="052b889",
        glob_include=['**/*.bas'],
        mid_process=mid_process,
        rename_overwrites={
            "FileCompress": "Compress",
        },

    ),

    # The following two projects are dependant on each other:
    Source(
        git_source="https://github.com/VBA-tools/VBA-JSON.git",
        git_rev="v2.3.1",
        glob_include=['**/JsonConverter.bas'],
        mid_process=mid_process,
        rename_overwrites={
            "JsonConverter": "Json", # bas file
            "Dictionary": "zJsonDict",
        },
    ),
    Source(
        git_source="https://github.com/VBA-tools/VBA-Dictionary.git",
        git_rev="757aea9",
        glob_include=['**/Dictionary.cls'],
        rename_overwrites={
            "Dictionary": "zJsonDict",
        }
    ),
    Source(
        git_source="https://github.com/sdkn104/VBA-CSV.git",
        git_rev="48d98d6",
        glob_include=['**/CSVUtils.bas'],
        mid_process=mid_process,
        rename_overwrites={
            "CSVUtils": "CsvUtils",
        },

    ),
    #Source(
    #    git_source="https://github.com/nylen/vba-common-library.git",
    #    git_rev="1e21b0d",
    #    glob_include=['**/VBALib_ExcelTable.cls', '**/VBALib_ExcelUtils.bas'],
    #    auto_bas_namespace=False,
    #    rename_overwrites={
    #        "VBALib_ExcelTable": "zListObject",
    #        "VBALib_ExcelUtils": "z__ListObject",
    #        "GetExcelTable": "zGetListObject"
    #    }
    #),
    Source(
        git_source="https://github.com/todar/VBA-Strings",
        git_rev="6d25dad",
        glob_include=["*.bas"],
        rename_overwrites={
            "StringFunctions":"StrUtils"
        }
    )
).run(
    output
)

# Turn off early bindings for "compress" module
cmpr = output.joinpath("z__Compress__.cls")
with cmpr.open("rb") as f:
    txt = f.read().replace(b"#Const EarlyBinding = True",
                           b"#Const EarlyBinding = False")
with cmpr.open("wb") as f:
    f.write(txt)


## Strip everything but the ExcelTable Factory
#lo = output.joinpath("z__ListObject.bas")
#with lo.open("rb") as f:
#    txt_lines = f.read().split(b"\r\n")
#
#with lo.open("wb") as f:
#    f.write(b"\r\n".join(txt_lines[:5]+txt_lines[400:435]))



"""
Possible VBA sources to choose from:

-- https://github.com/sancarn/stdVBA.git (looks promising)
-- https://github.com/ws-garcia/VBA-CSV-interface.git (Very nice!)
-- https://github.com/GustavBrock/VBA.Compress.git
-- https://github.com/VBA-tools/VBA-JSON.git
-- https://github.com/nylen/vba-common-library (VBALib_ExcelTable.cls)

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
https://github.com/x-vba/xlib
"""
