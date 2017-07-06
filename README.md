# cvstoexcel
Read a csv and writes a excel xlsx file merging columns with equal values.

This do not merge rows!!

This initial version only process columns from A to Z

Feel free to adapt for you needs.

Antonio Marques - July 2017 - aamarques@gmail.com


    usage: csvtoexcel.py [-h] [-i IFILENAME] [-o OFILENAME] [-d DELIMITER]
                         [-c LAST_COL] [-v]

    optional arguments:
      -h, --help         show this help message and exit
      -i IFILENAME       The input file to be parsed. REQUIRED
      -o OFILENAME       The output file without extension. Default is
                         outfile.xlsx
      -d DELIMITER       Delimite used in file. Default is space
      -c LAST_COL        Number of last column to merge. Default is 4
      -v, -V, --version  show program's version number and exit

    Ex: csvtoexcell.py -i example.csv -c 6 -d ';'

