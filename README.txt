Takes a directory of .txt, .csv and/or .xls(x) files and compiles a list of all field headers contained in the
files. This list is then used to generate a report of which files contain which field headers and which field headers
are contained in which files.

This is useful for determining Primary and Foreign Key relationships between multiple database compatible files. It will
also serve as the foundation for a future project I plan to implement, which will merge all files containing the same
primary key, so that a directory of files containing different data for the same record can be consolidated into fewer
files.
