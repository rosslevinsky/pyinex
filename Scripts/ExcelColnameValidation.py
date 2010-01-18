# $Id: ExcelColnameValidation.py 143 2009-09-08 15:50:33Z Ross $
#
# <license>
#
# This file is part of Pyinex, a project to embed python in Excel.
# 
# Copyright (c) 2009, Ross Levinsky
# All rights reserved.
#
# The Pyinex project is built using the xlw framework, found at 
# http://xlw.sourceforge.net
#
# Redistribution and use in source and binary forms, with or without 
# modification, are permitted provided that the following conditions are met:
#
#    * Redistributions of source code must retain the above copyright notice, 
#      this list of conditions and the following disclaimer.
#    * Redistributions in binary form must reproduce the above copyright 
#      notice, this list of conditions and the following disclaimer in the 
#      documentation and/or other materials provided with the distribution.
#    * Neither the name of Ross Levinsky nor the names of any other 
#      contributors may be used to endorse or promote products derived from 
#      this software without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.
#
#</license>*/

###############################################################################
#
# Conversion of Excel column number to column name is hairier than one might
# initially expect. The main issue is that there is no fixed, unchanging
# digit representing zero. Consider column name 'AA' - the 'A' on the right
# means "zero units of 1" and the 'A' on the left means '1 unit of 26." Worse,
# consider 'AAA', in which the middle 'A' now means "zero units of 26." 
# Essentially, 'A' means "zero" except when it's in the leftmost column.
#
# I arrived at a seemingly-correct iterative solution but had a hard time
# convincing myself that it was actually correct, so I turned to Python
# for testing. The algorithm is valid.
#
# This file contains test code to validate column name creation, number to
# name, and name to number routines.
#
###############################################################################

def GenerateColnames(numElements):

    left = [chr(ord('A') + i) for i in range(0,26)]
    right = left[:]
    left.insert(0,None)
    middle = left[:]

    # Flags indicating that we've started to use the more-left digits
    passedRight = False
    passedMiddle = False

    counter = 0
    retlist = []

    # Need to iterate over copies of these lists, as the for-in 
    # construct breaks when the iterated-over list is modified
    for i in left[:]:
        for j in middle[:]:
            for k in right:
                ov = ''
                if i: ov += i
                if j: ov += j
                ov += k
                
                # Remove the None vals from leftmost slots of left and middle
                # Flags enable us to do this pop() just once
                if k == 'Z': 
                    if passedRight == False:
                        middle.pop(0)
                    passedRight = True

                if j == 'Z': 
                    if passedMiddle == False:
                        left.pop(0)
                    passedMiddle = True

                retlist.append(ov)
                counter += 1
                if counter >= numElements:
                    return retlist

###############################################################################

# Text indexing is from the left of the string; reverse the input string
# so that the least-significant digit is always column 0

def NumFromName(name):

    t = name[::-1]

    if len(t) == 1:
        return (ord(t[0]) - ord('A'))

    if len(t) == 2:
        return  26* (ord(t[1]) - ord('A') +1) +  (ord(t[0]) - ord('A'))

    if len(t) == 3:
        return ((ord(t[2]) - ord('A') + 1) * 26 * 26 + 26) + \
            26* (ord(t[1]) - ord('A')) +  (ord(t[0]) - ord('A'))

###############################################################################
#
# Column name consists of three character 'digits': l, m, and r. 
# l and m may be empty. In the analysis below, I do character arithmetic
# as is common C/C++. The 'A' below is actually ord('A') in Python, and
# l, m, and r are also character code values. 'num' is the column number.
#
# ----
#
# (r -'A') covers [0,25] 
#
# Therefore: l = num + 'A'
#
# ----
#
# 26*(m -'A'+1) + (r-'A') covers [26, 26*26 + 25]
# 
# Extract digits with modular arthimetic:
# 
# (26*(m -'A'+1) + (r-'A')) % 26 = num % 26
# 
# Therefore: r = (num % 26) +  'A'
#            m = (num - (r - 'A'))/26 - 1 + 'A'
# 
# ----
#
# (26*26*(l-'A'+1) + 26) + 26*(m -'A') + (r-'A') 
# covers [26*27, 26*26*26 + 26 + 25*26 + 25 (= 26^3 + 26^2 + 25)
#
# Modular arithmetic again gives:
# 
# r = (num % 26) + 'A'
# 
# Thus:
#
# 26*(l-'A'+1) + (m -'A' + 1) = (num-(r-'A'))/26
# (m -'A') % 26 = (((num-(r-'A'))/26) - 1) % 26
#
# Thus:
# 
# m = ((((num-(r-'A'))/26) - 1) % 26) + 'A'
# l = ((((num-(r-'A'))/26) - 1  - (m-'A')) / 26) - 1 + 'A'

def NameFromNum(num):

    l = ''; m = ''
    a_ = ord('A')

    r = chr(a_ + num % 26) # True for all cases

    if num > 25 and num <= (26*26 + 25):
        m = chr((num - (ord(r) - a_))/26 - 1 + a_)

    elif num >= 26*27:
        m = ( (((num - (ord(r) - a_)) / 26) -1) % 26 ) + a_
        m = chr(m)

        l = num - (ord(r) - a_) - 26 * (ord(m) - a_) - 26
        l /= (26*26)
        l = l - 1 + a_
        l = chr(l)

    return l + m + r

###############################################################################
#
# Looking at the analysis of the individual digits above, we can see that
# a simple iterative algorithm is possible.
#
def IterativeNameFromNum(num):
    
    # This pulls out chars from the right; reverse string when done
    remainder = 0
    ov = ''
    while( True ):
        remainder = num % 26;
        ov += chr(ord('A') + remainder)
        if num < 26:
            break
        num = ((num - remainder) / 26) - 1;
        
    return ov[::-1]

###############################################################################

def main():

    maxCol = 16384
    correctNames = GenerateColnames(maxCol)
    formulaicNums  = map(NumFromName, correctNames)
    formulaicNames = map(NameFromNum,      range(0,maxCol))
    iterativeNames = map(IterativeNameFromNum, range(0,maxCol))

    if formulaicNums == range(0, maxCol):
        print "NumFromName is correct"

    if formulaicNames == correctNames:
        print "NameFromNum is correct"        

    if iterativeNames == correctNames:
        print "IterativeNameFromNum is correct"        

###############################################################################

if __name__ == '__main__': main()

###############################################################################


