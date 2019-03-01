import tkinter
from tkinter import Frame
from tkinter import Label
from tkinter import Entry
from tkinter import StringVar
from tkinter import messagebox

# so we can use sin, cos, pi, etc in our spreadsheet
import math
# we use this to make a deep copy of a dictionary
import copy

import dependencies

class Cell(Label):
	'''
	This class contains methods used by each cell
	in our spreadsheet grid
	'''

	def __init__(self,root,label):
		'''
		This function initializes a spreadsheet cell
		'''

		#Each cell will have a label, a string representing its value
		self.label = label
		self.prevString = ""
		#three attributes used for storing value in cell
		self.string = ""
		self.code = None
		self.value = None 
		#call the parent classes constructor for the label
		super().__init__(root, width=20, text=self.value, borderwidth=1,\
		 relief="solid", anchor=tkinter.W)
		self.root = root

	def updateCell(self):
		'''
		This function updates the cell and checks for errors
		'''
		
		# create these variables first as they will always need
		# to be used
		prevCode = self.code
		tempDependersOn = []
		prevDeps = copy.deepcopy(self.root.deps)

		if self.string != "":
			try:
				self.code = compile(self.string, self.label, 'eval')
			except Exception as e:
				self.code = prevCode
				self.string = self.prevString
				self.root.displayExceptionAlert(e)
				return False
			# First we check for the cyclic dependency
			co_names = self.code.co_names
			# here we find the actual nodes based on the names
			# in co_names
			for name in co_names:
				if self.root.isCellName(name):
					self.root.deps[self.label].add(name)
			try:
				tempDependersOn = dependencies.dependersOn(self.label, self.root.deps)
			except Exception as e:
				self.root.deps = prevDeps
				self.code = prevCode
				self.string = self.prevString
				self.root.displayExceptionAlert(e)
				return False
			# Now we try to evaluate the function
			try:
				self.value = eval(self.code, {}, self.root.symtab)
			except Exception as e:
				self.root.deps = prevDeps
				self.code = prevCode
				self.string = self.prevString
				self.root.displayExceptionAlert(e)
				return False
		else:
			self.code = None
			self.value = ""
			tempDependersOn = dependencies.dependersOn(self.label, self.root.deps)


		# we need to update dictionary now if successful
		# deep copy so once original is changed new one isnt affected
		prevSym = copy.deepcopy(self.root.symtab)
		self.root.symtab[self.label] = self.value

		# the last thing we have to do is check the nodes
		# that depend on this node and make sure there are no
		# errors when using the new value. This is somewhat 
		# inefficient as this same operation will be done later,
		# but it fits best with the logic of my solution
		for node in tempDependersOn:
			tempNode = self.root.findNode(node)
			try:
				dummy = eval(tempNode.code,{}, self.root.symtab)
			except Exception as e:
				self.root.symtab = prevSym
				self.root.deps = prevDeps
				self.code = prevCode
				self.string = self.prevString
				self.root.displayExceptionAlert(e)
				return False
		# if everything is good, we change
		# the label text
		self.configure(text=str(self.value))
		return True

class Spreadsheet(Frame):
	def __init__(self,root, nRows, nCols):
		'''
		This method initializes the spreadsheet class.
		'''

		# there cannot be more than 26 rows
		if nRows > 26:
			raise ValueError("Number of rows cannot exceed 26.")

		self.root = root
		super().__init__(root)
		self.nRows = nRows
		self.nCols = nCols

		# first we should put the labels on the spreadsheet for 
		# the rows/columns
		tempChar = ord('a')
		for i in range(1,nRows+1):
			tempLabel = Label(self, text=chr(tempChar))
			tempLabel.grid(row=i, column=0)
			tempChar+=1

		tempChar = 0
		for j in range(1,nCols+1):
			tempLabel = Label(self, text=str(tempChar))
			tempLabel.grid(row=0, column=j)
			tempChar+=1

		# we also need to initialize the symtab + deps dictionaries
		self.symtab = {}
		self.deps = {}

		# we need to add the math modules symbols
		#could add additional modules here
		self.addModuleSymbols(math)

		# Initialize the cell grid
		self.cellGrid = [[None for i in range(nCols)] \
			for j in range(nRows)]

		curRow, curCol = ord('a'), 0
		# We then initialize our cells
		for i in range(nRows):
			for j in range(nCols):
				tempLabel = chr(curRow) + str(curCol)
				self.cellGrid[i][j] = Cell(self, tempLabel)
				self.symtab[tempLabel] = ""
				self.deps[tempLabel] = set([])
				self.cellGrid[i][j].grid(row=i+1, column=j+1)
				self.cellGrid[i][j].bind('<Button-1>', lambda event, x=i, y=j: self.focus(x,y))
				# increment to the next column
				curCol += 1
			# reset column and go to next row
			curCol = 0
			curRow += 1

		# the focus cell will be the one with the yellow background
		# it starts as cell [0][0]
		self.focusPosition = (0,0)

		# now we initialize the focuslabel and focus entry
		self.focusLabel = Label(root,text="")
		self.entryContent = StringVar()
		self.focusEntry = Entry(root, textvariable=self.entryContent)
		# make it so after we hit enter, the grid is updated based on 
		# the contents of entry
		root.bind('<Return>', lambda event: self.enterPressed())
		# make it so after we hit tab, we move one to the left
		root.bind('<Tab>', lambda event: self.tabPressed())
		# makes it so that back-tab works too!
		root.bind('<ISO_Left_Tab>', lambda event: self.backTabPressed())

		
		# moves the focus left, right, up, down, respectively
		root.bind('<Left>', lambda event: self.moveLeft())
		root.bind('<Right>', lambda event: self.moveRight())
		root.bind('<Up>', lambda event: self.moveUp())
		root.bind('<Down>', lambda event: self.moveDown())


		#now we call updateFocus() for the first time
		self.updateFocus()

	def updateFocus(self):
		'''
		This functions updates the focus and frame GUI 
		whenever the focus is changed via keyboard/mouse
		'''

		# we go through the grid and update each cell based 
		# on current overall status
		for i in range(self.nRows):
			for j in range(self.nCols):
				# first we check to see if this is the focus label
				# if it is, we need to update it's background
				# and contents based on focus entry
				if (i,j) == self.focusPosition:
					self.cellGrid[i][j].configure(background="yellow")
					self.focusLabel.configure(text=\
						self.cellGrid[i][j].label)
					# update the focus entry to correspond with new cell
					self.focusEntry.delete(0, len(self.entryContent.get()))
					self.focusEntry.insert(0, self.cellGrid[i][j].string)
				else:
					self.cellGrid[i][j].configure(background="white")

	def updateGrid(self):
		'''
		This function updates the grid based on the cell that 
		gets a new value
		'''

		# we go through the list of dependencies for the cell at
		# x,y and update them all to the correct value

		# first we try to update the cell itself
		if self.cellGrid[self.focusPosition[0]][self.focusPosition[1]].updateCell():
			# update entry label
			self.focusEntry.delete(0, len(self.entryContent.get()))
			self.focusEntry.insert(0, self.cellGrid[self.focusPosition[0]][self.focusPosition[1]].string)
			# get the nodes that depend on this one
			tempDependersOn = dependencies.dependersOn(\
				self.cellGrid[self.focusPosition[0]][self.focusPosition[1]].label, self.deps)
			for node in tempDependersOn:
				tempNode = self.findNode(node)
				tempNode.updateCell()
		else:
			self.focusEntry.delete(0,len(self.entryContent.get()))
			self.focusEntry.insert(0, self.cellGrid[self.focusPosition[0]][self.focusPosition[1]].prevString)

	def focus(self, x,y):
		'''
		This function updates the focus attribute and calls
		to change the position of focus on GUI
		'''

		self.focusPosition = (x,y)
		self.updateFocus()

	def enterPressed(self):
		'''
		When enter is pressed, the focused cell's expression
		string is updated, as well as it's previous string.
		Then our updateGrid() function is called to perform the 
		calculations
		'''

		# update string on focused cell and previous string.
		self.cellGrid[self.focusPosition[0]][self.focusPosition[1]].prevString = self.cellGrid[self.focusPosition[0]][self.focusPosition[1]].string
		self.cellGrid[self.focusPosition[0]][self.focusPosition[1]].string = self.entryContent.get()
		self.updateGrid()

	def tabPressed(self):
		'''
		# functionality for tab presses
		# it will move to the right by default,
		# if at the end of the row, it will go to the next row
		# if it is at the bottom right corner, it will stop
		'''
		# per the specs, it needs to save the cell
		self.enterPressed()

		if self.focusPosition[0] != self.nRows -1 or self.focusPosition[1] != self.nCols - 1:
			if self.focusPosition[1] == self.nCols - 1:
				self.focus(self.focusPosition[0]+1,0)
			else:
				self.moveRight()

	def backTabPressed(self):
		'''
		Same as tab function above, but the opposite
		'''
		# per the specs, it needs to save the cell
		self.enterPressed()

		if self.focusPosition[0] != 0 or self.focusPosition[1] != 0:
			if self.focusPosition[1] == 0:
				self.focus(self.focusPosition[0]-1,self.nCols-1)
			else:
				self.moveLeft()

	# the following functions are used to move the focus with the 
	# arrow keys

	def moveLeft(self):
		# bounds check
		if self.focusPosition[1] > 0:
			self.focus(self.focusPosition[0],self.focusPosition[1] - 1)

	def moveRight(self):
		# bounds check
		if self.focusPosition[1] < self.nCols - 1:
			self.focus(self.focusPosition[0],self.focusPosition[1] + 1)

	def moveUp(self):
		# bounds check
		if self.focusPosition[0] > 0:
			self.focus(self.focusPosition[0] - 1,self.focusPosition[1])

	def moveDown(self):
		# bounds check
		if self.focusPosition[0] < self.nRows - 1:
			self.focus(self.focusPosition[0] + 1,self.focusPosition[1])


	def isCellName(self,name):
		'''
		This function determines if the string is a cell
		'''
		if name in self.symtab:
			return True
		else:
			return False


	def displayExceptionAlert(self,e):
		'''
		This function will display a message box with an error 
		message for the user to view
		'''
		messagebox.showerror(type(e).__name__, e)

	def addModuleSymbols(self,modulename):
		'''
		This function adds a module's functionality
		to the spreadsheet
		'''
		for value in modulename.__dict__:
			if value[:2] != "__":
				self.symtab[value]=modulename.__dict__[value]

		print("Imported symbols from module:", \
			modulename.__name__, "into the spreadsheet.")

	def findNode(self, nodeLabel):
		'''
		This helper function will find an actual cell in our
		cell grid based on it's label name
		'''

		tempChar = nodeLabel[0]
		tempInt = int(nodeLabel[1:])

		return self.cellGrid[ord(tempChar) - ord('a')][tempInt]




