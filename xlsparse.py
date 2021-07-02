
import re
import zipfile
from xml.etree import ElementTree as ET

def _dget(d,k):
	return d[k] if k in d else None

def _cell(s):
	a,b = re.match(r'^([A-Z]+)(\d+)$',s).groups()
	x = 0
	for c in a:
		x *= 26 # length of alphabet
		x += ord(c) - ord('A')
	# return (x,y) tuple for position.
	return (x,int(b)-1) # b-1 for zero-index position

class XLSFile:
	'''
	An object targeting an xlsx file, able to iterate over cells or rows.
	'''
	def __init__(self,target):
		self.target = target
	
	def sheets(self):
		'''
		Return a list of sheet indices.
		'''
		with zipfile.ZipFile(self.target) as z:
			r = []
			for s in z.namelist():
				m = re.search(r'/sheet(\d+).xml',s)
				if m:
					r.append(int(m.group(1)))
			return sorted(r)

	def rows(self,sheet):
		'''
		Iterate through sheet rows.
		
		Note that the rows do not necessarily begin at the spreadsheet's
		respective (0,0) position, but instead start at the earliest x and
		y, and end at the latest x and y.  The rows are of equal length and
		position relative to their own origin point, however.  If exact
		indices are desired, use cells().
		'''
		yt = {}
		minx = (1<<32) # arbitrary max value
		maxx = 0
		for pos,val in self.cells(sheet):
			x,y = pos
			minx = min(minx,x)
			maxx = max(maxx,x)
			if not y in yt:
				yt[y] = {}
			yt[y][x] = val
		for y in yt:
			row = [''] * (maxx-minx+1)
			for x in yt[y]:
				row[x-minx] = yt[y][x]
			yield row
	
	def cells(self,sheet):
		'''
		Iterate through sheet cells.
		'''
		with zipfile.ZipFile(self.target,'r') as z:
			# Get shared strings list/table, for reference in next part.
			vals = []
			with z.open('xl/sharedStrings.xml') as f:
				root = ET.parse(f).getroot()
				# Get namesapce, which need be included in .findall()s.
				namespace = re.sub(r'^({[^}]+})?.*',r'\1',root.tag)
				def findem(ns=''):
					# Iter through <si> nodes, adding values of <t> children.
					query = './/%ssi' % (ns,)
					for n in root.findall(query):
						r = list(map(lambda n:n.text, n.findall('.//%st' % (ns,))))
						vals.append(r)
					return len(vals)
				# Run against nodes without namespace, then with if that doesn't work.
				v = findem()
				if not v:
					findem(namespace)
			
			# Fill cell table with literal values or values dereferenced from vals.
			table = {}
			with z.open('xl/worksheets/sheet%i.xml' % (sheet,)) as f:
				root = ET.parse(f).getroot()
				# Get namesapce, again (though note there are two namespaces...?)
				namespace = re.sub(r'^({[^}]+})?.*',r'\1',root.tag)
				def findem(ns=''):
					count = 0
					for c in root.findall('.//%sc' % (ns,)):
						# Get pos from cell address.
						pos = _cell(c.attrib['r'])
						# Pull value from <v> node within <c>, based on cell type.
						t = c.attrib['t']
						val = c.find('.//%sv' % (namespace,)).text
						if t == 's':
							# String reference, so pull from vals.
							val = ''.join(vals[int(val)])
						table[pos] = val
						count = count + 1
					return count
				# Run against nodes without namespace, then with if that doesn't work.
				v = findem()
				if not v:
					findem(namespace)
		for r in table.items():
			yield r

if __name__ == '__main__':
	import sys
	
	def arg(k):
		return sys.argv[sys.argv.index(k)] if k in sys.argv else None
	def arg2(k):
		if k in sys.argv:
			return sys.argv[sys.argv.index(k)+1]
		else:
			return None
	def usage(m=''):
		if m:
			print(m)
		print('\n'.join([
			'',
			'usage:',
			'  xlsparse.py <target-xlsx file> <sheet #> [<output-type>]',
			'',
			'ouptut types:',
			'  -c,--csv   csv with quotes where commas are included in the value',
			'  -p,--pipe  pipe-delimited fields, no special character handling',
			'  -t,--tab   tab-delimited fields, no special characters',
			'  -s <s>,--sep <s>  delimit with contents of <s>',
			'',
			'other args:',
			'  -h,--help  output usage',
			''
		]))
	
	# Couple assertions.
	try:
		assert 2 < len(sys.argv) <= 5
	except AssertionError as e:
		usage()
		exit(1)
	
	# Interpret arguments.
	try:
		target = sys.argv[1]
		sheet = int(sys.argv[2])
		outtype = 'csv'
		delim = ','
		if   arg('-c') or arg('--csv'):
			outttype = 'csv'
			delim = ','
		elif arg('-p') or arg('--pipe'):
			outtype = 'pipe'
			delim = '|'
		elif arg('-t') or arg('--tab'):
			outtype = 'tab'
			delim = '\t'
		elif arg('-s') or arg('--sep'):
			outtype = 'sep'
			delim = arg2('-s') or arg2('--sep')
		elif arg('-h') or arg('--help'):
			usage()
			exit(0)
	except ValueError:
		usage('\nERROR: second arg (sheet) must be an integer')
		exit(1)
	except Exception as e:
		usage('Exception')
		raise e
	
	# Parse target file.
	xf = XLSFile(target)
	for r in xf.rows(int(sheet)):
		if outtype == 'csv':
			# In case of delims (,) within strings, enquote that particular value.
			def enquote(s):
				return ('"%s"' % (
					re.sub(r'"','\\"',s),) # escape quote chars if inserting quotes.
				) if delim in s else s
			r = list(map(enquote,r))
		try:
			print(delim.join(r))
		except UnicodeEncodeError:
			print(delim.join(map(lambda s:s.encode('utf-8'),r)))
	

