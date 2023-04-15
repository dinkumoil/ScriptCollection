# CombineFiles

This small script is able to merge the content of two files line by line into an output file.

To use the script you have to edit lines 3 to 5 in order to set the actual file names of the input and output files. If the two input files `InFile1` and `InFile2` have different numbers of lines, `InFile1` should be the one with more lines. In this case `OutFile` will contain empty lines for every missing line in `InFile2`.
