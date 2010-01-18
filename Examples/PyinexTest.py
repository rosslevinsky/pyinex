import sys

###############################################################################

watchMeGetIncremented = 4 # this is set when module is imported

def PlayWithGlobalVariable():
    global watchMeGetIncremented
    watchMeGetIncremented = watchMeGetIncremented + 1 
    return watchMeGetIncremented

###############################################################################

def TypesOfReturnValues(idx):
    retArray = [ [(2,2),('be', 1.0), 'me', 123.4, (1,2,3)],
                 ((2,2),('be', 1.0), 'me', 123.4, (1,2,3)),
                 ((2,2),('be', 1.0), 'me', 123.4, (1,2,3),(1,('too nested',))),
                 {'server': 'mpilgrim', 'database': 'master'},
                 [{'server': 'pilgrim', 'database': 'master'}],
                 None,
                 True,
                 False,
                 'me',
                 3.14,
                 666,
                 1234567891011121314 ]
    
    # idx comes in as float from Excel, but Python only takes int types 
    # for indices

    idx = int(idx) 
    if idx >= len(retArray):
        err = 'Index out of range: max val is ' + str(len(retArray)-1)
        print('TypesOfReturnValues: ' + err)
        return err

    results = retArray[idx]
    print('\nResults in Python:\n' + str(results))
    return results

###############################################################################
#
# Python 3.0 doesn't do automatic conversion of all list elements to strings
# when doing list sorts. We have to do this ourselves, but it can't be with
# the naive approach of passing a key=str param to sorted(). That just puts
# quotes around each of the top-level elements of the list to be sorted,
# which we can easily see gives undesirable behavior with nested lists:
#
# sorted( [1,'foo'], key=str)
# [1, 'foo']
# sorted( [[1],['foo']], key=str)
# [['foo'], [1]]
#
# To get sensible behavior, We have to go all the way down to the individual
# numeric elements, no matter how deeply they're nested, and turn them into 
# strings. This is done with the recursive function stringize(). We then see:
#
# sorted( [1,'foo'], key=stringize)
# [1, 'foo']
# sorted( [[1],['foo']], key=stringize)
# [[1], ['foo']]
#
# This works with Python 2.5, 2.6, and 3.0.

def stringize(x):
    if type(x) == list:
        # In 3.0, map returns an iterator (of type 'map') that can't
        # be sorted. Force it into an explicit list with the list c-tor.
        return list(map(stringize,x))
    elif type(x) == tuple:
        # Ditto with tuples
        return tuple(map(stringize,x))

    # str(None) == 'None', which doesn't sort where we expect. 
    #
    # Special-case it to return an empty string. 
    #
    # See http://boodebr.org/main/python/tourist/none-empty-nothing
    # to understand why we test with 'is' instead of ==

    elif x is None:
        return ''
    else:
        return str(x)

def SortList(unsortedList):
    sortedList = sorted(unsortedList, key=stringize)
    print('Unsorted: ' + str(unsortedList))
    print('Sorted  : ' + str(sortedList))
    return sortedList

###############################################################################

def CellWordcount(sentence):

    if type(sentence) != str: 
        return 0
    return len(sentence.split())

###############################################################################

# Global regexps; used so we can compile them only once

import re
pTime = re.compile('The Outstanding Public Debt as of (.* GMT)')

version = str.split(str.split(sys.version)[0], '.')
if int(version[0]) == 2:
    from HTMLParser import HTMLParser
    import urllib
else: 
    # Namespaces changed in 3.0
    from html.parser import HTMLParser
    import urllib
    import urllib.request

class DebtGetter(HTMLParser):
    def __init__(self, url):
        HTMLParser.__init__(self)
        self.timestamp = None
        self.debt = None
        if int(version[0]) == 2:
            self.req = urllib.urlopen(url)
        else:
            # Function location changed in 3.0
            self.req = urllib.request.urlopen(url)

        self.retcode = 200 # default version, as getcode() isn't present in 2.5
        if (( int(version[0]) == 2 and int(version[1]) > 5) or
            int(version[0]) == 3):
            self.retcode = self.req.getcode()
            if self.retcode != 200:
                print('Couldn\'t open site: ' + url + ' returned ' +  self.req.getcode())
                self.req.close()
                return

        self.feed(str(self.req.read()))

    def handle_starttag(self, tag, attrs):
        if self.debt is None and tag == 'img' and attrs:
            imgName = list(filter(lambda x : x[0] == 'src', attrs))
            if imgName is not None and  imgName[0][1] == 'debtiv.gif':
                self.debt = list(filter(lambda x : x[0] == 'alt', attrs))[0][1]
                self.debt = self.debt.replace(' ','')
                self.debt = self.debt.replace(',','')
                self.debt = self.debt.replace('$','')
                self.debt = float(self.debt) # it's a string until this operation

    def handle_data(self, data):
        if self.timestamp is None: # don't waste time if we've already found it
            match = pTime.search(data)
            if match:
                self.timestamp = match.group(1)
        
def USFederalDebt():
    debtURL = 'http://www.brillig.com/debt_clock/'
    # Does the parsing on construction
    parser = DebtGetter(debtURL)
    # Will close on destruction, I think, but that happens at an indefinite time,
    # so it's better to reclaim the connection resource now, deterministically
    parser.req.close()
    if parser.timestamp and parser.debt:
        # 2 x 1 tuple, so Excel displays it vertically stacked
        return ((parser.timestamp,), (parser.debt,))
    else:
        return 'Couldn\'t parse out debt (%s)' % self.retcode

###############################################################################
def Columnize( tup ):
    return [(x,) for x in tup]

def HasVarargs( firstParam, *allTheRest ):
    
    r1 = 'First params: ' + str(firstParam)
    r2 = str(len(allTheRest)) + ' varargs'
    r3 = 'Varargs: ' + ','.join([str(x) for x in allTheRest]) + ')'
    
    return Columnize( (r1, r2, r3) )

###############################################################################
#
# Trivial function exercised by TestHarness code
#

def TestHarnessFunc( i ):
    return (i, 2*i)   

###############################################################################  
           
 
