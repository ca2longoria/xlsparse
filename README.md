
A simple, dependency-free parser of xlsx files.

usage output:
```
usage:
  xlsparse.py <target-xlsx file> <sheet #> [<output-type>]

ouptut types:
  -c,--csv   csv with quotes where commas are included in the value
  -p,--pipe  pipe-delimited fields, no special character handling
  -t,--tab   tab-delimited fields, no special characters
  -s <s>,--sep <s>  delimit with contents of <s>

other args:
  -h,--help  output usage
```

For example...
```
python3 xlsparse.py TestSheet.xlsx 1 --tab
```
... the above will render a tab-delimited table.


