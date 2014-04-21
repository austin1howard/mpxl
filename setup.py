from setuptools import setup, find_packages
from mpxl import __version__

with open('longdesc.rst') as f:
    long_description = f.read()

setup(
	name='mpxl',
	version=__version__,
	packages=find_packages(),
	scripts = ['scripts/mpxl','scripts/mpxl_template'],
	install_requires = ['matplotlib','appscript','kaplot >= 0.9a'],
	author = 'Austin Howard',
	author_email = 'ahoward@utdallas.edu',
	url = 'http://github.com/austin1howard/mpxl',
	description = 'Matplotlib plotting tool for Microsoft Excel on OS X',
	long_description = long_description,
	classifiers = [
		'Development Status :: 4 - Beta',
		'Intended Audience :: Science/Research',
		'Intended Audience :: Financial and Insurance Industry',
		'Topic :: Scientific/Engineering :: Visualization',
		'Topic :: Office/Business :: Financial :: Spreadsheet',
		'Programming Language :: Python :: 2.6',
		'Programming Language :: Python :: 2.7',
		'Environment :: MacOS X',
		'Operating System :: MacOS :: MacOS X',
		'License :: OSI Approved :: MIT License'
	],
	data_files=[('share/mpxl', ['README.md', 'longdesc.rst',
                                'LICENSE']),]
)