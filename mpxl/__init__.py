# Library of helper functions for matplotlib excel interface
from string import lower,replace,split,strip,split
import kaplot
import kaplot.defaults as kd
from appscript import app
from appscript import k as k_app
from tempfile import NamedTemporaryFile
from subprocess import PIPE,Popen
from inspect import getargspec

__version__ = '1.4~beta1~km2'

_LAYERS = ['insettl', 'insettr', 'insetbl', 'insetbr', 'twinx', 'twiny']

_LAYER_SETTINGS = []
_LAYER_SETTINGS.append({'location' : 'upper left'})
_LAYER_SETTINGS.append({'location' : 'upper right'})
_LAYER_SETTINGS.append({'location' : 'lower left'})
_LAYER_SETTINGS.append({'location' : 'lower right'})
_LAYER_SETTINGS.append({'twin' : 'x'})
_LAYER_SETTINGS.append({'twin' : 'y'})

_IGNORE_ME		= '!' # if an entry starts with _IGNORE_ME all values of that index (row) are ignored

_LEGEND_LOCATIONS = ['upper right', 'upper left', 'lower left', 'lower right']

def _is_float(value):
	try:
		float(value)
		return True
	except:
		return False

def _convertToFloatOrBoolOrInt(x):
	"""if x can be converted to float or bool or int, do so and return the result"""
	try:
		if isinstance(val, (bool)):
			nv = bool(val)
		elif float(x) == int(x):
			nv = int(x)
		else:
			nv = float(x)
	except ValueError:
		nv = str(x)
		if lower(x) == 'true':
			nv = True 
		elif lower(x) == 'false':
			nv = False 
	return nv
	
def _splitEscaped(s,spl):
	"""splits `s` at `spl`, unless `spl` is preceeded by a backslash ("escaped")"""
	ret = s.replace('\\'+spl,'<>><<>').split(spl)
	ret = map(lambda x: x.replace('<>><<>',';'),ret)
	return ret

def _runKaplotFunction(k,fnName,fnArgs,fnKwargs):
	fn = getattr(k,fnName)
	argsNeeded = (getargspec(fn)[1] != None or len(getargspec(fn)[0]) > 1) # if the length is 1 it's just self
	if not argsNeeded and '=' in str(fnArgs):
		# most likely the user meant to pass kwargs in the second column
		fnKwargs = str(fnArgs)
	else:
		fnArgs = str(fnArgs)
		fnKwargs = str(fnKwargs)
		args = _splitEscaped(fnArgs,';')
		args = map(_convertToFloatOrBoolOrInt,args)
	kwargs = {}
	if fnKwargs != '':
		kwargsSplit = _splitEscaped(fnKwargs,';')
		for kwarg in kwargsSplit:
			key,value = kwarg.split('=')
			kwargs[key] = _convertToFloatOrBoolOrInt(value)
	if argsNeeded:
		fn(*args,**kwargs)
	else:
		fn(**kwargs)

def _get_path():
	applescriptCommand = 'POSIX path of (choose file name with prompt "Save PDF...")'
	p = Popen(['osascript','-e',applescriptCommand],stdout=PIPE)
	osxPath = p.communicate()[0].strip('\n')
	return osxPath


class ExcelSelection:
	def __init__(self):
		self._datasets = []
		self._layers = set(['main']) # it's a set so there are no duplicates
		self._layer_labels = {}
		self._layer_units = {}
		self._layer_colors = {}
		self.k = None # kaplot object
		self.showOnly = False #overridden if "show" keyword used
		self.pdf = False #overridden if 'pdf' keyword used

	def getSelection(self):
		"""
		Gets the active selection from excel using appscript module.
		Returns list of lists, and saves as self.selectionList

		If there's multiple selections, combine into one
		"""
		areas = app("Microsoft Excel").selection.areas.get()
		tmp_selectionList 	= areas.pop(0).value.get()
		self.selectionList 	= []
		for area in areas:
			tmp_selectionList = [row + area.value.get()[rowIndex] for rowIndex,row in enumerate(tmp_selectionList)]
		for i,selection in enumerate(tmp_selectionList):
			if not str(selection[0]).startswith(_IGNORE_ME):
				self.selectionList.append(selection)
		return self.selectionList

	def insertPlot(self):
		"""
		Inserts plot into new worksheet.
		"""
		newSheet = app(u'Microsoft Excel').active_workbook.make(at=app.active_workbook.end, new=k_app.worksheet)
		# get HFS style path
		applescriptCommand = 'return POSIX file "%s" as string' % self.ntf.name
		p = Popen(['osascript','-e',applescriptCommand],stdout=PIPE)
		osxPath = p.communicate()[0].strip('\n')
		newPic = app(u'Microsoft Excel').make(at=newSheet.beginning, new=k_app.picture, with_properties={k_app.file_name: osxPath, k_app.height: self.pixelSize[1], k_app.width: self.pixelSize[0]})
		self.ntf.close()

	def _determineRows(self):
		"""
		Determines the row layout of the spreadsheet
		Possible options are:
		- data only
		- label, data
		- label, unit, data
		- label, unit, legend label, data
		First line of data can optionally be a schema specifying the X,Y,Xerr,or Yerr columns,
		optionally with semicolon specifying the plotting layer to use, optionally with another semicolon
		separating kwargs to be passed to `add_plotdata`. If schema is not specified, assumes XYXYXY....

		All of this may optionally be proceeded with rows specifying plot options. Each row takes the form:
		`param`, `value` (multiple columns if needed), `kwargs` (multiple columns if needed). If param is "settings,"
		then specify either a semicolon separated list of "settings" in kaplot.defaults or separate them across columns.
		If param is "pdf", then second column should be a filename to save plot as in the default location.
		Otherwise, param must be a `set_` function in kaplot. (i.e., `set_title`) which will be run with the supplied arguments.)
		"""
		rowSpec = []
		currentRow = 0
		while True:
			col1 = self.selectionList[currentRow][0]
			# first check for params
			if all(cell == '' for cell in self.selectionList[currentRow]):
				rowSpec.append('blank')
			elif col1 == 'settings':
				rowSpec.append('settings')
			elif col1 == 'pdf':
				rowSpec.append('pdf')
				self.pdf = True
			elif col1 == 'show':
				self.showOnly = True # used to only show the plot
				rowSpec.append('show')
			elif type(col1) == type(u'') and col1.startswith('set_'):
				rowSpec.append('set_')
			elif type(col1) == type(u'') and col1.startswith('add_'):
				rowSpec.append('set_')
			elif _is_float(col1):
				# double check
				if _is_float(self.selectionList[currentRow+1][0]):
					# two rows of numbers in a row, probably onto data
					rowSpec.append('data')
					break # once we hit data, we're done
			elif type(col1) == type(u'') and lower(col1).startswith(('x;','y;','xerr;','yerr;')):
				rowSpec.append('schema')
			elif type(col1) == type(u'') and lower(col1) in ('x','y','xerr','yerr','_noshow_','_no_show_','_skip_'):
				rowSpec.append('schema')
			else:
				# if nothing above, must be in the label section.
				if rowSpec == []:
					rowSpec.append('label')
				elif rowSpec[-1] == 'label':
					rowSpec.append('units')
				elif rowSpec[-1] == 'units':
					rowSpec.append('legend')
				else:
					rowSpec.append('label')
			currentRow += 1

		# rowSpec returned
		return rowSpec

	def _standardizeSelection(self):
		"""
		Returns a selectionList which is in the standard format expected by extractParams.
		"""
		rowSpec = self._determineRows()
		selectionList = self.selectionList
		width = len(selectionList[rowSpec.index('data')]) # width of the data columns in excel

		# If settings specified, needs to be passed to kaplot.__init__
		try:
			settings_index = rowSpec.index('settings')
			settingsStrings = _splitEscaped(selectionList[settings_index][1],';')
			settings = map(lambda x: getattr(kd,x), settingsStrings)
		except ValueError:
			settings = None
		self.k = kaplot.kaplot(settings=settings)

		# If pdf specified, need to get filename and set variable
		try:
			pdf_index = rowSpec.index('pdf')
			self.pdf_filename = selectionList[pdf_index][1]
			if self.pdf_filename == '':
				# get from savebox
				self.pdf_filename = _get_path()
			if not self.pdf_filename.endswith('.pdf'):
				self.pdf_filename += '.pdf'
		except ValueError:
			pass

		# Check for any set_ rows. Also see if set_legend was explicitly specified.
		for i,r in enumerate(rowSpec):
			if r == 'set_':
				fnName = selectionList[i][0]
				fnArgs = selectionList[i][1]
				if lower(fnName) == 'set_legend':
					self.isLegend_specified = True
				else:
					self.isLegend_specified = False
				try:
					fnKwargs = selectionList[i][2]
				except IndexError:
					fnKwargs = u'' # only two columns selected
				_runKaplotFunction(self.k, fnName, fnArgs, fnKwargs)

		self.isLegend = 'legend' in rowSpec

		# assemble the rest of the things
		standardSelectionList = []

		for rowName in ['label','units','legend']:
			if rowName in rowSpec:
				standardSelectionList.append(selectionList[rowSpec.index(rowName)])
			else:
				standardSelectionList.append([''] * width)

		if 'schema' in rowSpec:
			schemaRow = selectionList[rowSpec.index('schema')]
		else:
			schemaRow = ['X','Y'] * (width/2)
		standardSelectionList.append(schemaRow)

		# Add data
		dataList = selectionList[rowSpec.index('data'):]
		standardSelectionList += dataList
		self.standardSelectionList = standardSelectionList

		# Clear _noshow_ and _skip_ columns
		for col,colType in reversed(list(enumerate(schemaRow))):
			if colType in ['_no_show_','_noshow_','_skip_']:
				# remove that column
				for row in standardSelectionList:
					del row[col]
		return standardSelectionList

	def extractParams(self):
		"""
		Extracts the header information from the selection. Going down the rows step by step, we look for the following information.
		If it's found, advance to the next row and go to the next step, otherwise just go to the next step for the same row.

		Schema:
		Each cell contains one of the following:
			X, Y, Xerr, Yerr
		to specify the type of data in that column. Columns pair in the same way as Origin. For example, X | Y | Y | Yerr
		would produce 2 curves, both with the same x data, and the second with error bars on y.
		Optionally, this can be followed by a semicolon (;) and then followed by one of the following layer names:

			insetTL, insetTR, insetBL, insetBR (for inset plot in one of 4 locations)
			twinx,twiny (for plots on a second pair of y or x axes, respectively.)

		If left out, we assume the data is on the "main" layer. For example, the following is a valid schema:

			1 | 2 | 3          | 4          | 5       | 6       | 7 | 8       | 9
			X | Y | X;insetTL  | Y;insetTL  | X;twinx | Y;twinx | X | Y;twiny | Y

		which will produce a plot like:
			
			               5,6
			    +-----------------------+
			    |   +--+                |
			    |3,4|  |                |
			1,2 |   +--+                |
			5,6 |   3,4                 | 7,8
			7,9 |                       |
			    |                       |
			    |                       |
			    +-----------------------+
			               1,2
			               7,8
			               7,9
			
		The numbers show the xy data pairing, and which axes that pairing will be plotted against.
		The layer name may be followed by another semicolon (;) and a semicolon separated list of kwargs: e.g., X;main;lw=10;marker=o
		self.isLegend = False
		Data:
			Data is collected in the corresponding MPLDataSet object.
		"""
		selectionList = self._standardizeSelection()
		
		self.labels = selectionList[0]
		self.units = selectionList[1]
		self.legend = selectionList[2]
		self.schema = selectionList[3]
		self.dataStartRow = 4
		self.processSchema()

	def processSchema(self):
		"""
		Parses through the schema list to determine the different layers present and datasets to use
		"""
		schema = self.schema
		xCol = None # The column representing the current set of x values. Each time an "X" is encountered this will update.
		xErr = None # The column representing the current x error. None if there's no error.
		yCol = None # Ditto for y
		yErr = None # ditto for y
		skip = 0 # Number of columns to skip at start of loop. This is used if finding information (errors) on the next column
		for c,s in enumerate(schema):
			# Skip needed columns
			if skip > 0:
				skip -= 1
				continue

			# Check for new X
			if lower(s).startswith('x') and not lower(s).startswith('xerr'):
				xCol = c
				# Is next column errors?
				if len(schema) > c + 1 and lower(schema[c+1]).startswith('xerr'):
					xErr = c + 1
					skip += 1
				else:
					xErr = None

			# If not X, should be a Y column
			if lower(s).startswith('y') and not lower(s).startswith('yerr'):
				yCol = c
				# Is next column errors?
				if len(schema) > c + 1 and lower(schema[c+1]).startswith('yerr'):
					yErr = c + 1
					skip += 1
				else:
					yErr = None
				# what layer should it be in?
				if len(s) > 1 and s[1] == ';':
					layerInfo = split(s[2:],';',1)
					layer = layerInfo[0]
					if len(layerInfo) > 1:
						# get the kwargs
						kwargsString = layerInfo[1]
						kwargs = {}
						for kwarg in _splitEscaped(kwargsString,';'):
							key,value = kwarg.split('=')
							kwargs[key] = _convertToFloatOrBoolOrInt(value)
					else:
						kwargs = {}
				else:
					layer = 'main'
					kwargs = {}
				# This is a complete dataset
				self._datasets.append(MPLDataSet(self,xCol,xErr,yCol,yErr,lower(layer),kwargs))
				self._layers.add(lower(layer))
				if lower(layer) not in self._layer_labels:
					self._layer_labels[lower(layer)] = (self.labels[xCol],self.labels[yCol])
				if lower(layer) not in self._layer_units:
					self._layer_units[lower(layer)] = (self.units[xCol] or None,self.units[yCol] or None)
				if lower(layer) not in self._layer_colors:
					# see if color specified:
					if u'color' in kwargs.keys():
						self._layer_colors[lower(layer)] = kwargs['color']
					else:
						self._layer_colors[lower(layer)] = None


	def makePlot(self,show=False):
		"""
		Makes plot in matplotlib using kaplot extension, and saves to temporary file. If show == True, displays plot (for debugging only)
		"""
		k = self.k # TODO change all k's to self.k's
		# Add all the layers (except 'main')
		# self._layers.remove('main')
		for lname in self._layers:
			if lname == 'main':
				continue # don't need to explicitly add 'main'
			layer = _LAYERS.index(lname) # this will error out if wrong layer name was input
			k.add_layer(lname,**(_LAYER_SETTINGS[layer]))
		# Add all the data
		for dataset in self._datasets:
			kwargs = dataset.kwargs
			if self.isLegend:
				k.add_plotdata(x=dataset.xData,y=dataset.yData,xerr=dataset.xErr,yerr=dataset.yErr,name=dataset.layer,label=(dataset.label or '_nolegend_'), **kwargs)
			else:
				k.add_plotdata(x=dataset.xData,y=dataset.yData,xerr=dataset.xErr,yerr=dataset.yErr,name=dataset.layer, **kwargs)
		for i,lname in enumerate(self._layers):
			kwargs = {}
			if lname != 'main' and self._layer_colors[lname]:
				kwargs['color'] = self._layer_colors[lname]
			# Allow for _ to be used as lack of units
			if self._layer_units[lname][0] == '_':
				xlab = None
			else:
				xlab =self._layer_units[lname][0]
			if self._layer_units[lname][1] == '_':
				ylab = None
			else:
				ylab =self._layer_units[lname][1]
			k.set_xlabel(lab=self._layer_labels[lname][0],unit=xlab,name=lname, **kwargs)
			k.set_ylabel(lab=self._layer_labels[lname][1],unit=ylab,name=lname, **kwargs)
			if self.isLegend and self.isLegend_specified == False:
				k.set_legend(True,loc=_LEGEND_LOCATIONS[i],name=lname)
		# calculate plot size in pixels
		dpi = k.SAVEFIG_SETTINGS['dpi']
		width = k.SAVEFIG_SETTINGS['width']
		height = k.SAVEFIG_SETTINGS['height']
		self.pixelSize = (width * 72 , height * 72)
		k.makePlot()
		if show or self.showOnly:
			k.showMe()
		elif self.pdf:
			k.saveMe(self.pdf_filename)
		else:
			self.ntf = NamedTemporaryFile(delete=False,suffix='.png')
			k.saveMe(self.ntf.name,dpi=dpi)

class MPLDataSet:
	"""
	This contains a single dataset (X,Y,Xerr,Yerr) that is passed to kaplot.add_data
	"""
	def __init__(self,selection,xCol,xErr,yCol,yErr,layer,kwargs):
		"""
		Extracts data from selectionList, and adds to MPLLayer
		"""
		dataList = selection.standardSelectionList[selection.dataStartRow:]
		dataList = map(list, zip(*dataList))
		self._xData = dataList[xCol]
		self._yData = dataList[yCol]
		if xErr:
			self._xErr = dataList[xErr]
		else:
			self._xErr = None
		if yErr:
			self._yErr = dataList[yErr]
		else:
			self._yErr = None
		## ALL DATA SETS HAVE BEEN BUILT
		## CLEANUP DATA SETS
		# remove blank lines
		while self._xData[-1] == '':
			self._xData.pop()
			self._yData.pop()
			if self._xErr is not None:
				self._xErr.pop()
			if self._yErr is not None:
				self._yErr.pop()
		# remove bad entries (non-float)
		self.xData = []
		self.yData = []
		if self._xErr is None:
			self.xErr = None
		else:
			self.xErr = []
		if self._yErr is None:
			self.yErr = None
		else:
			self.yErr = []
		for i,x in enumerate(self._xData):
			y = self._yData[i]
			if type(x) == type(0.0) and type(y) == type(0.0):
				self.xData.append(x)
				self.yData.append(y)
				if self.yErr is not None:
					if type(self._yErr[i]) == type(0.0):
						self.yErr.append(self._yErr[i])
					else:
						self.yErr.append(0.0)
				if self.xErr is not None:
					if type(self._xErr) == type(0.0):
						self.xErr.append(self._xErr[i])
					else:
						self.xErr.append(0.0)
		## END OF CLEANUP
		self.layer = layer
		self.kwargs = kwargs
		if selection.isLegend:
			self.label = selection.legend[yCol]
		else:
			self.label = None
