# excel-shred
Utility for converting excel workbook with multiple sheets, to multiple json and or csv data sets

Pass in one or more paths to search for all Excel files.  
Each located file will be parsed into a json data set, csv data set or both.
The new file(s) will be in a folder of the same name as the original file


```

  Open an Excel workbook, and convert all sheets to json datasets 
  
  outdir: output directory for files :param format: the output format
  input_dirs: one or more directory paths containing excel workbooks

  Examples:

  excel-shred input_dir_a [input_dir_b]

  excel-shred -f csv input_dir_a [input_dir_b]

  excel-shred -f csv -o .\output input_dir_a [input_dir_b]

Options:
  -f, --format [json|csv|all]
  -o, --outdir PATH
  -h, --help                   Show this message and exit.
```  

# Format Options

* all (default) --output both json and csv
* json --output only json format
* csv --output only csv format
