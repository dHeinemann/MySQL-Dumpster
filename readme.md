**MySQL Dumpster** is a small VBScript to automatically backup all MySQL databases on a Windows server.

It will:

* fetch a list of databases in the MySQL instance
* ignore any blacklisted databases
* backup & compress the remainders
* move the compressed backups to a specified directory

#Usage
MySQL Dumpster takes no command line parameters. All configuration is
done by editing the script itself.

#Requirements
MySQL Dumpster requires the command-line version of 7-Zip (7za.exe), which is used to compress the backups. It can be downloaded from [http://www.7-zip.org/download.html](http://www.7-zip.org/download.html).

#License
MySQL Dumpster is published under the MIT License.

    Copyright (c) 2012 David Heinemann

    Permission is hereby granted, free of charge, to any person obtaining a
    copy of this software and associated documentation files (the
    "Software"), to deal in the Software without restriction, including
    without limitation the rights to use, copy, modify, merge, publish,
    distribute, sublicense, and/or sell copies of the Software, and to
    permit persons to whom the Software is furnished to do so, subject to
    the following conditions:

    The above copyright notice and this permission notice shall be included
    in all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
    CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
    TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
    SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
