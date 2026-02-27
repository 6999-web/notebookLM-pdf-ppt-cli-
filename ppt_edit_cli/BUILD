load("@rules_python//python:defs.bzl", "py_binary", "py_library")

py_library(
    name = "ppt_edit_lib",
    srcs = [
        "__init__.py",
        "engine.py",
        "utils.py",
    ],
    deps = [
        # Dependencies would normally be defined in WORKSPACE/requirements.txt
    ],
)

py_binary(
    name = "ppt-edit-cli",
    srcs = ["main.py"],
    main = "main.py",
    deps = [
        ":ppt_edit_lib",
    ],
)
