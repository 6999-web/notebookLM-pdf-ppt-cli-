load("@rules_python//python:defs.bzl", "py_binary", "py_library")

py_library(
    name = "notebook_editor_lib",
    srcs = ["__init__.py", "main.py"],
    visibility = ["//visibility:public"],
)

py_binary(
    name = "notebook_editor",
    srcs = ["main.py"],
    main = "main.py",
    deps = [":notebook_editor_lib"],
)
