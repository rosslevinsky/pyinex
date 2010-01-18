###############################################################################
#
# Script does a fresh svn checkout of the tree, copies in binaries that the
# user has already compiled, and zips up the tree. 
#
# A better script would also do the build, run tests, and clean up the build
# for packaging, but this will do (for now).
#

import sys, glob, os, stat, shutil, tempfile, zipfile

###############################################################################
#
# Good example of how to use the zipfile module, with lots of pitfalls explained:
#
# http://bytes.com/groups/python/492744-script-make-windows-xp-readable-zip-file
#

def main():

    if len(sys.argv) != 3:
        usage()
        exit()
        
    version = sys.argv[1]
    xllSrc = sys.argv[2]

    # Temp dir to hold the code and binaries
    td = tempfile.mkdtemp(dir='.');
    versionedPyinex = 'Pyinex-' + version

    # Get fresh code checkout (requires cygwin to be installed)
    print("Checking out clean copy of Pyinex code...\n")
    os.system('svn co file://"/z:/documents/Repository/Pyinex/trunk" ' + td + '/' + versionedPyinex)

    # Copy in binaries
    print("\nCopying binaries...\n")
    xllDest = os.path.normpath(td + '\\' + versionedPyinex + '\Bin\\')
    for f in glob.glob(xllSrc + '\*.xll'):
        print f
        f = os.path.normpath(f)
        shutil.copy(f, xllDest)

    # Create zipfile. Setting 'w' mode overwites any existing file with this name.
    print("\nCreating zip file...\n")
    zf = zipfile.ZipFile(versionedPyinex + '.zip', 'w')
    
    cwd = os.getcwd()
    os.chdir(td)

    for dirpath,dirs,files in os.walk(''): # This starts the walk at the CWD
        for a_file in files:
            a_path = os.path.join(dirpath,a_file)

            # Don't want to ship any SVN control dirs
            if '.svn' not in a_path:
                zf.write(a_path)

            # SVN control dirs are not writable; need to make them so to be able to 
            # delete the tmpdir-rooted tree
            os.chmod(a_path, (stat.S_IWRITE | stat.S_IREAD))

    # Clean up
    print("\nCleaning up...\n")
    os.chdir(cwd)
    zf.close()
    shutil.rmtree(td)

###############################################################################

def usage():
    u = "usage: ' + os.path.basename(sys.argv[0]) + \
<Pyinex version number> <path to binaries>"
    print(u)
    
###############################################################################

if __name__ == '__main__':
    main()

###############################################################################
