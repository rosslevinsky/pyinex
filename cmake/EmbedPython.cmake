# Writes a C++ header holding one Python file's text, so the core compiles it in
# (docs/architecture.md, "The core's Python files"). CMakeLists.txt runs it at build time:
#
#   cmake -DSOURCE=<file.py> -DOUTPUT=<header> -DNAME=<identifier> -P EmbedPython.cmake
#
# The text is written as an array of characters, ending with a null, rather than as a
# string literal, which Microsoft's compiler limits to 16,380 characters a piece.

file(READ "${SOURCE}" hex HEX)
# Sixteen bytes to a line, then each byte as a character literal.
string(REPEAT "[0-9a-f][0-9a-f]" 16 line)
string(REGEX REPLACE "(${line})" "\\1\n" hex "${hex}")
string(REGEX REPLACE "([0-9a-f][0-9a-f])" "'\\\\x\\1'," bytes "${hex}")
file(WRITE "${OUTPUT}"
  "// Generated from ${SOURCE} by cmake/EmbedPython.cmake. Do not edit.\n"
  "#pragma once\n"
  "namespace pyinex::core::python {\n"
  "inline constexpr char ${NAME}[] = {\n${bytes}'\\0'};\n"
  "}\n")
