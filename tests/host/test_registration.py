"""xlAutoOpen registers exactly the four documented functions (docs/architecture.md, "Excel
functions"), through the table-driven helper."""

ARGUMENTS = "file,function," + ",".join(f"arg{i}" for i in range(1, 31))


def test_registers_the_documented_functions(host):
    registered = {r.name: r for r in host.registrations()}
    assert set(registered) == {"PyCall", "PyConsole", "PyReloadOnChange", "PyInfo"}
    expected = {
        "PyCall": ("Q" * 33, ARGUMENTS, 32),
        "PyConsole": ("QQQQQQQ", "show,clear,x,y,width,height", 6),
        "PyReloadOnChange": ("QQ", "on", 1),
        "PyInfo": ("QQ", "item", 1),
    }
    for name, (type_, arguments, help_count) in expected.items():
        r = registered[name]
        assert (r.procedure, r.type, r.arguments) == (name, type_, arguments), name
        assert r.category == "Pyinex", name
        assert r.help, name
        assert len(r.argument_help) == help_count and all(r.argument_help), name


def test_no_function_is_thread_safe(host):
    # Pyinex runs only on Excel's main thread (docs/architecture.md, "Threading").
    assert all("$" not in r.type for r in host.registrations())


def test_pycall_argument_text_fits_excels_limit(host):
    # Excel truncates a function's argument text beyond 255 characters.
    (pycall,) = [r for r in host.registrations() if r.name == "PyCall"]
    assert len(pycall.arguments) <= 255
    assert all(len(h) <= 255 for h in pycall.argument_help)
