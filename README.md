## xls2json

```bash
pip3 install xls2json
```

```bash
xls2json [--perentry] [--persheet] xls_input [output_path]
```

- **--perentry** Read XLS file and write a JSON file per entry of XLS file.
- **--persheet** Read XLS file and write a JSON file per sheet of XLS file.
- **default** Read XLS file and write to single JSON file.

## TODO

- auto detect cell type and write accordingly.
- handle XLS file with no table headers.
- handle non unique sheet names.
