SQLDumpSplitter3
================

Welcome to the SQLDumpSplitter3.
This is a small tool to split up large MySQL dumps into small, independently executable SQL files.
It is useful in cases where an upload limit is in place in web administration tools for exampel.
This little program is a revival of a very old program of
[mine](https://philiplb.de/sqldumpsplitter2/2016/01/25/a-glimpse-from-the-past-the-sql-dump-splitter/).

But this time, it is far more useful:

- it is based on a minimal, use case tailored SQL parser
- it splits also INSERTs with many value tuples over the files
- it is cross platform

So, happy splitting and report any bugs, you might encounter!

## Command line interface

This tool also comes along with a CLI. Here is how it is used:

```bash
SQLDumpSplitter3 split --file FILE --size SIZE --unit UNIT
```

* split: tells the program to execute a SQL split and not start the GUI. Shortcut: s
* --file FILE: the FILE to split. Must be given. Shortcut: -f FILE
* --size SIZE: the maximum file size, min. 100 KB. Default: 2. Shortcut: -s
* --unit UNIT: the unit of the maximum file size, one of "KB", "MB", "GB". Default: "KB" Shortcut: -u

Example calls:

```bash
# simple split with 2 MB files
SQLDumpSplitter3 split --file dump.sql

# the same using the shortcuts
SQLDumpSplitter3 s -f dump.sql

# this results in max 512 KB files
SQLDumpSplitter3 split --file dump.sql --size 512 --unit KB

# this results in max 512 KB files using the shortcuts
SQLDumpSplitter3 s -f dump.sql -s 512 -u KB
```

## Changelog

### 0.9.3

released 2020-10-16

- added the command line interface
- tracking USE statements and repeating them in the splitted files
- fixed a bug that the UTF8 BOM didn't get recognized as whitespace and so the very first query was always seen as general query and not an INSERT for example
- fixed a bug caused by the standard library buffered reader where the Peek function influences the Read position and so single "/" where removed from the resulting SQL

### 0.9.2

released 2018-12-22

- fixed a bug where the SQL got split incorrectly if string literals like `N'Foo'` were used

### 0.9.1

released 2018-09-02

- optimized the performance of the splitting by a factor of 4
- added the original base filename to the target directory and split file names

### 0.9

released 2018-08-30

- initial release

The application is [CC BY NC ND licensed](https://creativecommons.org/licenses/by-nc-nd/4.0/).
The license text should be delivered with the application as LICENSE.txt.

The application icon is based on the icon "project-diagram" by [Font Awesome](https://fontawesome.com).

Cheers
Philip
Philip@PhiliplB.de
https://philiplb.de
