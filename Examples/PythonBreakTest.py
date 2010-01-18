import time

# Note that no "import pyinex" statement is required (though you may
# add one if you like)

###############################################################################
#					
# The default for pyinex.Break() is to NOT clear the break; a single press of 
# the escape key will make all calls to	pyinex.Break() (in a single calculation
# cycle, from different cells) return True. If you pass a True into the 
# function, it will clear the break, and each cell that calls it will have a 
# chance to continue calculating.					
					
def LongLoop(id, count, clearBreak):

    # Remember - vars from Excel come in as floats, but range() expects
    # an int. Cast it to suppress warning.

    for i in range(0,int(count)):
        print(id + ' ' + str(i))
        time.sleep(1)
        if pyinex.Break(clearBreak):
            print('Excel escape key pressed')
            return id + ' aborted'


    return id + ' completed'

###############################################################################
