# Python scripts for Excel editing:

A collection of a few scripts I've thrown together to do some 
tedious work on larg(ish) excel files

Most of the replacements are made with replacement LUT files, 
with examples shown in the [dicts](./dicts) folder 

All of the scripts are meant to be called in the command line,
and generally support naming of columns by either the title row (row 1)
or Excel's column designation.

Eg: if Column `'A'` has a name `'ID Number'` in its `'A1'` cell, you can refer to it as either method using the 
`-d` | `--direct_col` flag to use excel's direct column names. The default is to use the column names to provide 
tolerance of shifting column orders. 


Most of them support a `-h` | `--help` option that prints the usage pattern. 

## notes:
The scripts are still pretty limited in some(**Most**) aspects, 
but if I need to use these more I will probably extend 
them and incorporate a shell aspect to be able to
perform multiple actions much quicker.
  
