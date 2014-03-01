# Library of helper functions for matplotlib excel interface
from string import lower

class ExcelSelection:
	def getSelection():
		"""
		Gets the active selection from excel using appscript module.
		Returns list of lists.
		"""
		return

	def extractParams(selectionList):
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
				twinx,twiny,twinxy (for plots on a second pair of x or y or both axes, respectively.)
				(For insets, all of the associated columns must have this after them.)
			For example, the following is a valid schema:
				1 | 2 | 3          | 4          | 5       | 6       | 7 | 8       | 9
				X | Y | X;inset:TL | Y;inset:TL | X;twinx | Y;twinx | X | Y;twiny | Y
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
		currentRow = 0
		# Title
		if lower(selectionList[currentRow][0]) == 'title':
			self.title = list(selectionList[currentRow][1])
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

class MPLDataSet:
	"""
	This contains a single dataset (X,Y,Xerr,Yerr) that is passed to kaplot.add_data
	"""
	def __init__(selectionList,xCol,yCol,xErr,yErr,layer):
		"""
		Extracts data from selectionList, and adds to MPLLayer
		"""

class MPLLayer:
	"""
	This contains the collection of MPLPlots that form a single MPL layer. There may be more than one MPLPlotGroup in the final plot.
	For example, if the schema is X,Y,Yerr,Y,X,Y, there are two plot groups: one has the first X, two Y plots, one with error bars.
	The second group has the second X and Y pair.

	Essentially, while scanning the schema from left to right, a new plot group is started under the following conditions:
	- X is encountered
	- 
	"""
	self.plotType = 'line'
	self.xData = []
	self.xErr = []
	self.yData = [[]]
	self.yErr = [[]]
