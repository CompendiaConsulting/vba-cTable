# cTable Regression Tests

This is an initial, self-contained test suite for the cTable VBA library. It uses in-memory fixtures; it does not depend on external sample workbooks.

## Installation and execution

1. Create a **test-only** macro-enabled Excel workbook (`.xlsm`).
2. Import the current cTable classes and `JsonConverter.bas` into that workbook. Preserve exported class attributes, especially `cLineageRegistry`'s `VB_PredeclaredId = True`.
3. Import `Test_cTableRegression.bas` as a standard module using **File > Import File** in the VBA editor.
4. In the VBA editor, select **Debug > Compile VBAProject**.
5. Open the Immediate Window (`Ctrl+G` in the VBE) and execute `RunAllTests` from the Macros dialog or Immediate Window.
6. Read `PASS` and `FAIL` lines and the final summary. Failed cases print VBA error numbers and descriptions; after all cases, the suite raises an error if any failed.

The suite tests import/accessors, filtering, the corrected Null policy, sort copy/in-place semantics, Append name alignment, concat, left merge/indicator, projection, clone independence, duplicate removal, grouped differences, and a basic lineage smoke test.

## Expected filtering requirement

The `TestMissingPolicy` test requires the **corrected** implementation: `nullBehavior:="exclude"` omits missing cells in both positive and inverse filters; `"include"` retains missing cells in both. If the October 7, 2026 plugin baseline is still in use without the filtering correction, failures in this test are expected and useful. Do not modify the test to accommodate the known bug.

## Limitations

- These tests have been reviewed against the plugin's bundled October 7, 2026 cTable API, **not compiled or executed in Excel**.
- The current repository source may differ from the bundled snapshot. Review failing signatures against the current `.cls` implementation.
- This is a starter suite, not comprehensive certification of the library.
- It does not exercise external workbook/CSV/JSON I/O, Excel for Mac or Windows interoperability, performance, deep lineage propagation through joins/sorts, or extensive edge cases.
- Avoid checking real financial data into the public repository. Use synthetic fixtures.

## Suggested repository layout

```text
tests/
├── README.md
└── Test_cTableRegression.bas
```

Maintain the tests whenever public behavior changes. A failing test should prompt inspection of both the code and the expected result, not automatic weakening of the assertion.
