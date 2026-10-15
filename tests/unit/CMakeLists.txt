# _pyinex_testing.pyd, the add-in's value conversion as a Python extension module, which
# tests/unit loads (tests/unit/conftest.py). It uses the stable ABI, as the core does, so
# one build serves every Python the tests run under.

add_library(_pyinex_testing MODULE testing/pyinex_testing.cpp)
set_target_properties(_pyinex_testing PROPERTIES SUFFIX ".pyd")
target_link_libraries(_pyinex_testing PRIVATE pyinex_python)
pyinex_defaults(_pyinex_testing)
