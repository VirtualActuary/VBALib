from zebra_vba_packager import Config, Source
from locate import this_dir

output = this_dir().joinpath("output")

def mid_process(source):
    for pth in source.temp_transformed.rglob("*.bas"):
        with pth.open("rb") as f:
            txt_lines = f.read().replace(b"\r", b"").decode("utf-8").split("\n")

        do_overwrite = False
        for i, line in enumerate(txt_lines):
            if line.strip().startswith("Public"):
                if line.strip().split()[2].lower() == "as":
                    do_overwrite = True
                    ii = line.lower().find("public")
                    line = line[0:ii] + "Private" + line[ii+len("public"):]
                    txt_lines[i] = line

        if do_overwrite:
            with pth.open("wb") as f:
                print(pth)
                f.write(("\r\n".join(txt_lines)).encode("utf-8"))


Config(
    Source(
        git_source="https://github.com/ws-garcia/VBA-CSV-interface.git",
        git_rev="v3.1.0",
        glob_include=['**/src/*.cls'],
        rename_overwrites={
            "ECPArrayList": "zCSVArray", # useful
            "CSVinterface": "z__CSV__",  # useful
            "ECPTextStream": "z__CSVTextStream",
            "parserConfig": "z__CSVParserConf",
        }
    ),
    Source(
        git_source="https://github.com/GustavBrock/VBA.Compress.git",
        git_rev="052b889",
        glob_include=['**/*.bas'],
        rename_overwrites={
            "FileCompress": "Compress",
        },
        mid_process=mid_process,
    ),

    # The following two projects are dependant on each other:
    Source(
        git_source="https://github.com/VBA-tools/VBA-JSON.git",
        git_rev="v2.3.1",
        glob_include=['**/JsonConverter.bas'],
        rename_overwrites={
            "JsonConverter": "JSON", # bas file
            "Dictionary": "zJSONDict",
        },
        mid_process=mid_process
    ),
    Source(
        git_source="https://github.com/VBA-tools/VBA-Dictionary.git",
        git_rev="757aea9",
        glob_include=['**/Dictionary.cls'],
        rename_overwrites={
            "Dictionary": "zJSONDict",
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
