# Library of helper functions for matplotlib excel interface
from string import lower,replace
import kaplot
from appscript import app,k
from tempfile import NamedTemporaryFile

__version__ = '0.1'

_LAYERS = ['insettl', 'insettr', 'insetbl', 'insetbr', 'twinx', 'twiny']

_LAYER_SETTINGS = []
_LAYER_SETTINGS.append({'location' : 'upper left'})
_LAYER_SETTINGS.append({'location' : 'upper right'})
_LAYER_SETTINGS.append({'location' : 'lower left'})
_LAYER_SETTINGS.append({'location' : 'lower right'})
_LAYER_SETTINGS.append({'twin' : 'x'})
_LAYER_SETTINGS.append({'twin' : 'y'})

_LEGEND_LOCATIONS = ['upper right', 'upper left', 'lower left', 'lower right']

"""app(u'Microsoft Excel').active_workbook.make(at=app.active_workbook.end, new=k.worksheet)"""

class ExcelSelection:
	def __init__(self):
		self._datasets = []
		self._layers = set(['main']) # it's a set so there are no duplicates
		self._layer_labels = {}
		self._layer_units = {}
		self.k = None # kaplot object

	def getSelection(self):
		"""
		Gets the active selection from excel using appscript module.
		Returns list of lists, and saves as self.selectionList
		"""
		self.selectionList = app(u'Microsoft Excel').selection.value.get()
		return self.selectionList

	def insertPlot(self):
		"""
		Inserts plot into new worksheet.
		"""
		newSheet = app(u'Microsoft Excel').active_workbook.make(at=app.active_workbook.end, new=k.worksheet)
		osxPath = 'Mavericks:' + self.ntf.name.replace('/',':')
		newPic = app(u'Microsoft Excel').make(at=newSheet.beginning, new=k.picture, with_properties={k.file_name: osxPath, k.height: 480, k.width: 640})
		self.ntf.close()

	def extractParams(self):
		"""
		Extracts the header information from the selection. Going down the rows step by step, we look for the following information.
		If it's found, advance to the next row and go to the next step, otherwise just go to the next step for the same row.

		1) Title (If A1 is the word title, B1 is the title of the plot. Otherwise there's no title, and this row is thought of as row 2.)

		2) TODO: Plot Style (If A2 is the word style, B2 is one of: line, scatter, hist, etc. If not specified assume line plot,
			and treat this row as row 3.)

		3) Label (required)

		4) Units (required)

		5) Legend entries:
			If row starts with a single "X", assume it is for the schema and skip the legend. Otherwise the items here will be put
			into a legend should there be more than one Y per X.

		6) Schema:
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
		
		Data:
			Data is collected in the corresponding MPLDataSet object.
		"""
		selectionList = self.selectionList
		currentRow = 0
		# Title
		if lower(selectionList[currentRow][0]) == 'title':
			self.title = selectionList[currentRow][1]
			self.isTitle = True
			currentRow += 1
		else:
			self.isTitle = False

		# Label and Units
		self.labels = selectionList[currentRow]
		currentRow += 1
		self.units = selectionList[currentRow]
		currentRow += 1

		# Legend, if it exists
		if lower(selectionList[currentRow][0]) != 'x':
			# There is a legend row
			self.legend = selectionList[currentRow]
			self.isLegend = True
			currentRow += 1
		else:
			self.isLegend = False

		# Schema
		self.schema = selectionList[currentRow]
		self.dataStartRow = currentRow + 1
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
					# there's more!
					layer = s[2:]
				else:
					layer = 'main'
				# This is a complete dataset
				self._datasets.append(MPLDataSet(self,xCol,xErr,yCol,yErr,lower(layer)))
				self._layers.add(lower(layer))
				self._layer_labels[lower(layer)] = (self.labels[xCol],self.labels[yCol])
				self._layer_units[lower(layer)] = (self.units[xCol],self.units[yCol])


	def makePlot(self):
		"""
		Makes plot in matplotlib using kaplot extension, and saves to temporary file.
		"""
		k = kaplot.kaplot()
		# Add all the layers (except 'main')
		# self._layers.remove('main')
		for lname in self._layers:
			if lname == 'main':
				continue # don't need to explicitly add 'main'
			layer = _LAYERS.index(lname) # this will error out if wrong layer name was input
			k.add_layer(lname,**(_LAYER_SETTINGS[layer]))
		# Add all the data
		for dataset in self._datasets:
			if self.isLegend:
				k.add_plotdata(x=dataset.xData,y=dataset.yData,xerr=dataset.xErr,yerr=dataset.yErr,name=dataset.layer,label=dataset.label)
			else:
				k.add_plotdata(x=dataset.xData,y=dataset.yData,xerr=dataset.xErr,yerr=dataset.yErr,name=dataset.layer)
		# And the rest of the stuff
		if self.isTitle:
			k.set_title(self.title)
		for i,lname in enumerate(self._layers):
			k.set_xlabel(lab=self._layer_labels[lname][0],unit=self._layer_units[lname][0],name=lname)
			k.set_ylabel(lab=self._layer_labels[lname][1],unit=self._layer_units[lname][1],name=lname)
			if self.isLegend:
				k.set_legend(True,loc=_LEGEND_LOCATIONS[i],name=lname)
		k.makePlot()
		# k.showMe()
		self.ntf = NamedTemporaryFile(delete=False,suffix='.png')
		k.saveMe(self.ntf.name,height=6,width=8,dpi=80)

class MPLDataSet:
	"""
	This contains a single dataset (X,Y,Xerr,Yerr) that is passed to kaplot.add_data
	"""
	def __init__(self,selection,xCol,xErr,yCol,yErr,layer):
		"""
		Extracts data from selectionList, and adds to MPLLayer
		"""
		dataList = selection.selectionList[selection.dataStartRow:]
		dataList = map(list, zip(*dataList))
		self.xData = dataList[xCol]
		self.yData = dataList[yCol]
		if xErr:
			self.xErr = dataList[xErr]
		else:
			self.xErr = None
		if yErr:
			self.yErr = dataList[yErr]
		else:
			self.yErr = None
		self.layer = layer
		if selection.isLegend:
			self.label = selection.legend[yCol]
		else:
			self.label = None
