def test_imports_smoke():
    # Basic smoke test so CI/devs catch missing deps or broken module moves.
    import pandas  # noqa: F401
    import openpyxl  # noqa: F401

    from parsers import PARSER_MAP  # noqa: F401
    from reporting.excel_report import generate_report  # noqa: F401
