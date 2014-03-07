from setuptools import setup, find_packages

setup(
	name='mpxl',
	version='0.1',
	packages=find_packages(),
	scripts = ['scripts/mpxl'],
	install_requires = ['matplotlib','appscript']
)