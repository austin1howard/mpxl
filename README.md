mplxl
=====

OS X package which grabs selection from Microsoft Excel, and plots with `kaplot` extension to `matplotlib`. Then inserts plot into a new worksheet in Excel.

Includes a script which is installed into python's `bin` folder. This can be run from command line, or could be included into an Automator script to attach to keyboard shortcut.

Requires `kaplot` and `appscript` modules, both available from [pypi](http://pypi.python.org).