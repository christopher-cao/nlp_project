#!/usr/bin/env python
#
# LexisNexis article splitter
# 
# PRELIMINARY VERSION -- NEEDS TESTING
# author: silvio.petriconi@unibocconi.it
#
# required packages: lxml, xlsxwriter, pandas


from lxml import html  
from lxml.etree import tostring
from itertools import chain
import os, sys
import re  # regular expressions
import pandas as pd
import json

from UserList import UserList

import argparse, sys, os
from argparse import ArgumentParser


def stringify_children(node):  
	"""Converts all children nodes of node to HTML string (with tags)"""
	parts = ([node.text] + 
    	list(chain(*([c.text, tostring(c), c.tail] for c in node.getchildren()))) + 
    	[node.tail])
	return ''.join(filter(None, parts))


def find_parent_with_tag(node, tag, root = None):
	"""Returns a parent node of node with tag (if exists), else None; searches upwards until reaching root"""
	if root == None:
		tree = node.getroottree()
		root = tree.getroot()
	node = node.getparent()
	while node.tag != tag and node != root:
		node = node.getparent()
	return node if node.tag == tag else None



class LexisNexisDocuments(UserList):
	def __init__(self, html_infile):

		super(self.__class__, self).__init__()  # in Python 3 this becomes super(self).__init__()
		html_doc = html_infile.read()

		# remove the HTML comments around the doc tags before parsing
		purged_html = re.sub("""\<![ \r\n\t]*(-- Hide XML([^\-]|[\r\n]|-[^\-])*--[ \r\n\t]*)\>""",\
			lambda x : x.group(0)[34:-4], html_doc)

		self.root = html.fromstring(purged_html)  # save the lxml tree
		self._construct_selectors()  # construct the right xpath selectors based on CSS properties

		for doc in self.root.xpath('//docfull'):
			document = {}
			docnode = doc.getparent()
			if docnode.tag != 'doc':
				print "Error: Some <docfull> has no parent <doc> node. Aborting."
				sys.exit(-1)
			document['ID'] = docnode.get('number')
			document['HTML'] = stringify_children(doc)
			print "Processing document number {id}".format(id=document['ID'])

			document.update(self._extract_details(doc))   	# add details\
			document.update(self._extract_text(doc))		# extract text 
			self.append(document) # finally, insert document in list

	def write_xlsx(self, outfilename):
		"""Writes the documents as Excel file."""
		df = pd.DataFrame.from_records(self)
		writer = pd.ExcelWriter(outfilename, engine='xlsxwriter')
		df.to_excel(writer, 'Sheet 1')
		writer.save()
		
	def write_json(self, f):
		"""Writes the documents as JSON."""
		json.dump(self._json_rep(), f)

	def get_json_str(self):
		"""Returns the documents as JSON."""
		return json.dumps(self._json_rep())

	def _json_rep(self):
		def dict_rm(keylist, dic):
			result = dic.copy()
			for key in keylist:
				result.pop(key, None)
			return result 
		return dict([(a["ID"], dict_rm(["ID","HTML"], a)) for a in self.data])  # map document ID to document		

	def _construct_selectors(self):
		# defines the right selectors: extract dynamic formatting class from CSS style headers
		stylenode = next(s for s in self.root.xpath("//head/style"))
		css = stylenode.text_content().strip()
		assert css.startswith("<!--") and css.endswith("-->"), "CSS expected in comment tags!"
		css = css[4:-3] # cut off the comment tags
		classre = re.compile(r"\.(c[0-9]+)\s+(\{.+\})")
		cssdict = {}  # empty dictionary

		for line in css.split("\n"):
			classmatch = classre.match(line)
			if classmatch != None: # if relevant css line found
				value, key = classmatch.group(1,2)
				cssdict[key] = value

		try:
			c1 = cssdict["{ text-align: center; margin-top: 0em; margin-bottom: 0em; }"]
			c2 = cssdict["{ font-family: 'Times New Roman'; font-size: 10pt; font-style: normal; font-weight: normal; color: #000000; text-decoration: none; }"]
			c3 = cssdict["{ text-align: center; margin-left: 13%; margin-right: 13%; }"]
			c4 = cssdict["{ text-align: left; }"]
			c5 = cssdict["{ text-align: left; margin-top: 0em; margin-bottom: 0em; }"]
			c6 = cssdict["{ font-family: 'Times New Roman'; font-size: 14pt; font-style: normal; font-weight: bold; color: #000000; text-decoration: none; }"]
			c7 = cssdict["{ font-family: 'Times New Roman'; font-size: 10pt; font-style: normal; font-weight: bold; color: #000000; text-decoration: none; }"]
			c8 = cssdict["{ text-align: left; margin-top: 1em; margin-bottom: 0em; }"]
		except KeyError:
			print "Error: CSS header does not define essential styles. Can't continue."
			sys.exit(-1)

		self.title_selector = './/div[@class="{c4}"]/p[@class="{c5}"]/span[@class="{c6}"]/parent::*'.format(c4=c4,c5=c5,c6=c6)
		self.headers_selector = './/div[@class="{c3}"]/p[@class="{c1}"]'.format(c3=c3,c1=c1,c2=c2)
		self.detail_selector = './/div[@class="{c4}"]/p[@class="{c5}"]/span[@class="{c7}"]/parent::*'.format(c4=c4,c5=c5,c7=c7)
		self.detailkey_selector = './span[@class="{c7}"]'.format(c7=c7)
		self.detailvalue_selector = './span[@class="{c2}"]'.format(c2=c2)
		self.text_selector = './/div[@class="{c4}"]/p[@class="{c8}"]/parent::*'.format(c4=c4,c8=c8)
		self.alternative_text_selector = './div[@class="{c4}"]/p[@class="{c5}"]/span[@class="{c2}"]/parent::p/parent::div'.format(c4=c4,c5=c5,c2=c2)


	def _extract_details(self, node):  # extracts the metadata details from a docfull node
		result = {}
		re_month = re.compile(r"(january|february|march|april|may|june|july|august|september|october|november|december)")

		for detail in node.xpath(self.detail_selector):
			key = ''.join(t.text_content() for t in detail.xpath(self.detailkey_selector)).strip().upper().rstrip(':')
			value = ''.join(t.text_content() for t in detail.xpath(self.detailvalue_selector)).upper()
			result[key] = value
			#print key, "|", value

		# then, extract title
		result['TITLE'] = ''.join([t.text_content() for t in node.xpath(self.title_selector)])
		idx = find_parent_with_tag(node, 'doc').get('number')

		# next, extract headers and footers, erasing unbreakable spaces
		headers_footers = [line.text_content().replace(u'\xa0',' ') for line in node.xpath(self.headers_selector)]

		if len(headers_footers) not in set([2,3]):
			print "WARNING: Unexpected number of {n} headers and/or footers in article {artnr}: {hf}".format(artnr=idx,n=len(headers_footers),hf=str(headers_footers))
		
		if len(headers_footers) < 3:  # something is missing: source, date, or copyright. Find which is which.
			report_date = -1
			copyright = -1
			for i in range(len(headers_footers)):
				if re_month.search(headers_footers[i].lower()):
					report_date = i  # seems like this is a date field
				if headers_footers[i].lower().startswith("copyright"):
					copyright = i
			
			result['REPORT_DATE'] = headers_footers[report_date] if report_date >= 0 else ''

			if copyright < 0:  # in case of missing copyright, try to find an alternatively formatted one
				candidates = [[i for i in node.xpath('.//div/p/span[starts-with(text(),"Copyright")]')],[i for i in node.xpath('.//div[starts-with(text(),"COPYRIGHT")]')]]
				for cnodes in candidates:
					if len(cnodes) > 0: # if found, save it
						cstr = cnodes[-1].text_content().replace(u'\xa0', ' ') # remove nonbreakable spaces
						print "WARNING: rescued mis-formatted copyright footer in article {n}: {s}".format(n=idx,s=cstr)
						headers_footers.append(cstr)
						copyright = len(headers_footers)-1 # last entry is copyright
						break

			result['COPYRIGHT'] = headers_footers[copyright] if copyright >= 0 else ''

			if report_date == 0 or copyright == 0:  # if first element was identified as copyright or date, then there's no source
				result['SOURCE'] = "" 
			else:
				result['SOURCE'] = headers_footers[0]
		elif len(headers_footers) == 3:
			result['SOURCE'], result['REPORT_DATE'], result['COPYRIGHT'] = headers_footers
		else:
			print headers_footers
			print "WARNING!!!!!!!! Length of headers and footers is {n} (and > 3) in article {artnr}! Headers and copyright WILL be misread.".format(artnr=idx, n=len(headers_footers))
			print "Extracted data must be fixed manually. Search for 'FIXME'."
			result['COPYRIGHT'] = result['SOURCE'] = "FIXME"
			report_date = -1
			for i in range(len(headers_footers)):
				if re_month.search(headers_footers[i].lower()):
					report_date = i  # seems like this is a date field
			result['REPORT_DATE'] = headers_footers[report_date] if report_date > 0 else "FIXME"

		# check sanity. Note that Chinese government agency articles did not have copyright.
		if (not result['COPYRIGHT'].lower().startswith('copyright')) and ('xinhua' not in result['COPYRIGHT'].lower()):
			print "Potentially misclassified copyright found in article {artnr}: {hf} => {cpy}".format(artnr=idx, hf=str(headers_footers),cpy=result['COPYRIGHT'])

		if not re_month.search(result['REPORT_DATE'].lower()):
			print "No valid date in article {artnr}".format(artnr=idx)

		return result

	def _extract_text(self, node):
		"""Extracts all text from a docfull node"""
		text_nodes = self._find_text_nodes(node.getparent())
		# now, within our text, make sure that <BR> yields '\n':
		for node in text_nodes:
			for br in node.xpath(".//br"):
				br.tail = "\n" + br.tail if br.tail else "\n"
		
		fulltext = ''.join(t.text_content() for t in text_nodes)
		finaltext = fulltext.replace(u'\xa0', ' ') # replace nonbreakable spaces
		finaltext = finaltext.replace('\n', ' ') # replace newlines
		return {'FULLTEXT':finaltext}

	def _find_text_nodes(self, docnode):
		"""Returns for a given doc node all the lxml tree nodes representing actual text."""
		def argmax(iterable, keyfun = lambda x:x):
			return max(enumerate(iterable), key=lambda x: keyfun(x[1]))[0]
		
		assert docnode.tag=='doc', "<doc> tag expected in parameter node."
		idx = docnode.get('number')

		docfullnode = next(m for m in docnode.xpath('.//docfull')) # there should only be one docfull child node

		text_nodes = [m for m in docfullnode.xpath(self.text_selector)]
		if len(text_nodes) == 0:
			print "Missing text in article {artnr}, trying to auto-recover by considering alternative formatting.".format(artnr=idx)
			print "Selecting with xpath " + self.alternative_text_selector
			tnodes = [m for m in docfullnode.xpath(self.alternative_text_selector)] # find div node
			text_nodes = [tnodes[argmax(tnodes, keyfun=lambda x : len(x.text_content()))]] # choose longest div node

			if len(text_nodes) > 0: 
				print "Success! Extracted text: => " + ''.join(t.text_content() for t in text_nodes)

		assert len(text_nodes) > 0, "Error: no text found in article {artnr}".format(artnr=idx)
		allnodes = [n for n in docfullnode.iterchildren()]

		try:
			pos_start = allnodes.index(text_nodes[0])
		except ValueError:
			raise RuntimeError("Error: text node not recovered from xpath tree. Article nr. {artnr}".format(artnr=idx))

		pos = pos_start
		while allnodes[pos].tag != "br":  # continue searching until you find the next <BR>
			pos += 1

		return allnodes[pos_start:pos]   


def main():
	""" MAIN PROGRAM """

	parser = ArgumentParser(description="LexisNexis text splitter")
	parser.add_argument('-i','--infile', type=argparse.FileType('r'), 
                      required=True)

	args = parser.parse_args()
	#html_doc = args.infile.read()  # read the HTML document
	path, filename = os.path.split(args.infile.name)
	filenameroot, ext = os.path.splitext(filename)
	outfilename = os.path.join(path, filenameroot + '.xlsx')

	docs = LexisNexisDocuments(args.infile)
	print "Writing Excel "+outfilename
	docs.write_xlsx(outfilename)

	with open('test_output.txt','w') as f:
		for doc in docs:
			f.write('-----------------------------{ID}----------------------------\n\n'.format(ID=doc['ID']) + doc['FULLTEXT'].encode('utf-8') + '\n')

if __name__ == "__main__":
    main()

