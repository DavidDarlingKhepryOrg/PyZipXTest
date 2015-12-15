import io
import logging
import os
import platform
import py7zlib
import subprocess
import zipfile

zipFileNamesList = []
zipFileNamesList.append('./testFiles/test47z.7z')
zipFileNamesList.append('./testFiles/test4zipx.zipx')
zipFileNamesList.append('./testFiles/test4zip.zip')
# zipFileNamesList.append('./testFiles/CID-Synonym-filtered.zipx')
# zipFileNamesList.append('./testFiles/CID-Synonym-filtered.7z')

tmpZipPath = './OutputFiles'

def main():
    
    errStr = None
     
    if platform.system() != 'Windows':
        p7zipPgmPath = '7z'
    else:
        p7zipPgmPath = os.path.abspath('/Program Files/7-Zip/7z.exe')
    
    tmpZipPathExpanded = getPathExpanded(tmpZipPath,
                                         parentPath=None,
                                         pgmLogger=logging)
    print('tmpZipPathExpanded: %s' % tmpZipPathExpanded)
    
    for zipFileName in zipFileNamesList:
        
        zipFileNameExpanded = getPathExpanded(zipFileName,
                                          parentPath = None,
                                          pgmLogger=logging)
        print('')
        print('---------------------------------------------')
        print('zipFileNameExpanded: %s' % zipFileNameExpanded)
    
        tmpZipPathExpanded, errStr = zip2dirX(zipFileNameExpanded,
                                              tmpZipPathExpanded,
                                              useSubProcessCall=True,
                                              p7zipPgmPath=p7zipPgmPath,
                                              isTestMode=True,
                                              pgmLogger=logging)

    return                    

    
# Unzip a 7Z, ZIP, or ZIPX archives to a target folder

def zip2dirX(zipFileName,
             tmpZipPath,
             useSubProcessCall=True,
             p7zipPgmPath='7z' if platform.system() != 'Windows' else os.path.abspath('/Program Files/7-Zip/7z.exe'),
             isTestMode=False,
             pgmLogger=None):
        
    errStr = None
    
    tmpZipPathExpanded = getPathExpanded(tmpZipPath,
                                         parentPath=None,
                                         pgmLogger=pgmLogger)
    
    if isTestMode:
        print('tmpZipPathExpanded: %s' % tmpZipPathExpanded)
    
    zipFileNameExpanded = getPathExpanded(zipFileName,
                                          parentPath = None,
                                          pgmLogger=pgmLogger)
    
    if isTestMode:
        print('')
        print('---------------------------------------------')
        print('zipFileNameExpanded: %s' % zipFileNameExpanded)
    
    if os.path.exists(zipFileNameExpanded):
        if os.path.isfile(zipFileNameExpanded):
            pieces = os.path.splitext(zipFileNameExpanded)
            if len(pieces) == 2:
                if pieces[1].lower() in ['.7z']:
                    try:
                        with io.open(zipFileNameExpanded, 'rb') as fp:
                            archive7z = py7zlib.Archive7z(fp)
                            # obtain the archive's manifest of file names
                            fileNameManifestList = archive7z.getnames() 
                            # if it's a 7-Zip file
                            if len(fileNameManifestList) > 0:
                                if isTestMode:
                                    for fileName in fileNameManifestList:
                                        print('7-Zip manifest file name: %s' % fileName)
                            else:
                                errStr = 'Specified 7Z archive file "%s" has an empty manifest' % zipFileNameExpanded
                                pgmLogger.error(errStr)
                    except Exception as e:
                        errStr = 'Specified 7Z archive file "%s" was not recognized as a valid "7ZIP" format' % zipFileNameExpanded
                        pgmLogger.error(errStr)
                        errStr = str(e)
                        pgmLogger.error(errStr)
                    if errStr == None:
                        if useSubProcessCall:
                            try:
                                subprocess.call([p7zipPgmPath, 'e', '-y', '-o%s' % tmpZipPathExpanded, zipFileNameExpanded], shell=False)
                            except Exception as e:
                                errStr = 'Decompression of ZIP archive file "%s" via subprocess call of 7-Zip program failed' % zipFileNameExpanded
                                pgmLogger.error(errStr)
                                errStr = str(e)
                                pgmLogger.error(errStr)
                        else:
                            try:
                                with io.open(zipFileNameExpanded, 'rb') as fp:
                                    archive7z = py7zlib.Archive7z(fp)
                                    # obtain the archive's manifest of file names
                                    fileNameManifestList = archive7z.getnames()
                                    # extract each file from the manifest 
                                    for fileName in fileNameManifestList:
                                        try:
                                            outfilename = os.path.join(tmpZipPathExpanded, fileName)
                                            outdir = os.path.dirname(outfilename)
                                            if not os.path.exists(outdir):
                                                os.makedirs(outdir)
                                            with io.open(outfilename, 'wb') as outfile:
                                                outfile.write(archive7z.getmember(fileName).read())
                                            if isTestMode:
                                                print ('outFileName: "%s"' % outfilename)
                                        except Exception as e:
                                            errStr = str('Expansion of manifest file "%s" within 7Z archive file "%s" failed' % (fileName, zipFileNameExpanded))
                                            pgmLogger.error(errStr)
                                            errStr = str(e)
                                            pgmLogger.error(errStr)
                            except Exception as e:
                                errStr = 'Decompression of 7Z archive file "%s" via "py7zlib" package failed' % zipFileNameExpanded
                                pgmLogger.error()
                                errStr = str(e)
                                pgmLogger.error(errStr)
                elif pieces[1].lower() in ['.zipx']:
                    try:
                        if zipfile.is_zipfile(zipFileNameExpanded):
                            # obtain the archive's manifest of file names
                            fileNameManifestList = zipfile.ZipFile(zipFileNameExpanded).namelist()
                            if len(fileNameManifestList) > 0:
                                if isTestMode:
                                    for fileName in fileNameManifestList:
                                        print('ZIP Manifest file name: %s' % fileName)
                            else:
                                errStr = 'Specified ZIPX archive file "%s" has an empty manifest' % zipFileNameExpanded
                                pgmLogger.error(errStr)
                        else:
                            errStr = 'Specified ZIPX archive file "%s" was not recognized as a valid "ZIPX" format' % zipFileNameExpanded
                            pgmLogger.error(errStr)
                    except Exception as e:
                        errStr = 'Specified ZIPX archive file "%s" was not recognized as a valid "ZIPX" format' % zipFileNameExpanded
                        pgmLogger.error(errStr)
                        errStr = str(e)
                        pgmLogger.error(errStr)
                    if errStr == None:
                        try:
                            # decompress the archive to a temporary folder
                            subprocess.call([p7zipPgmPath, 'e', '-y', '-o%s' % tmpZipPathExpanded, zipFileNameExpanded], shell=False)
                        except Exception as e:
                            errStr = str(e)
                            pgmLogger.error('Decompression of ZIPX archive file via subprocess call of 7-Zip program failed')
                            pgmLogger.error(errStr)
                elif pieces[1].lower() in ['.zip']:
                    try:
                        if zipfile.is_zipfile(zipFileNameExpanded):
                            # obtain the archive's manifest of file names
                            fileNameManifestList = zipfile.ZipFile(zipFileNameExpanded).namelist()
                            if len(fileNameManifestList) > 0:
                                if isTestMode:
                                    for fileName in fileNameManifestList:
                                        print('ZIP Manifest file name: %s' % fileName)
                            else:
                                errStr = 'Specified ZIP archive file "%s" has an empty manifest' % zipFileNameExpanded
                                pgmLogger.error(errStr)
                        else:
                            errStr = 'Specified ZIP archive file "%s" was not recognized as a valid "ZIP" format' % zipFileNameExpanded
                            pgmLogger.error(errStr)
                    except Exception as e:
                        errStr = 'Specified ZIP archive file "%s" was not recognized as a valid "ZIP" format' % zipFileNameExpanded
                        pgmLogger.error(errStr)
                        errStr = str(e)
                        pgmLogger.error(errStr)
                    if errStr == None:
                        if useSubProcessCall:
                            try:
                                subprocess.call([p7zipPgmPath, 'e', '-y', '-o%s' % tmpZipPathExpanded, zipFileNameExpanded], shell=False)
                            except Exception as e:
                                errStr = 'Decompression of ZIP archive file "%s" via subprocess call of 7-Zip program failed' % zipFileNameExpanded
                                pgmLogger.error(errStr)
                                errStr = str(e)
                                pgmLogger.error(errStr)
                        else:
                            try:
                                # extract all of the files/folders
                                # within the archive file to the target folder
                                with zipfile.ZipFile(zipFileNameExpanded, "r") as z:
                                    z.extractall(tmpZipPathExpanded)
                            except Exception as e:
                                errStr = 'Decompression of ZIP archive file "%s" via "zipfile" package failed' % zipFileNameExpanded
                                pgmLogger.error(errStr)
                                errStr = str(e)
                                pgmLogger.error(errStr)
                else:
                    errStr = 'Nominal archive file "%s" is not in an ARAPCD-supported archival format (.7z, .zip, .zipx)'
                    pgmLogger.error(errStr)
            else:
                errStr = 'Nominal archive file "%s" does not have an extension' % zipFileNameExpanded
                pgmLogger.error(errStr)
        else:
            errStr = 'Specified archive file "%s" was not recognized as a valid file' % zipFileNameExpanded
            pgmLogger.error(errStr)
    else:
        errStr = 'Specified archive file: "%s" does not exist' % zipFileNameExpanded
        pgmLogger.error(errStr)
            
    return tmpZipPathExpanded, errStr                    


# =============================================================================    
# Create a folder if it does not already exist
# =============================================================================    

def createFolderIfNotExist(folderNameExpanded,
                           containsFileName=False,
                           pgmLogger=None):
    
    errStr = None
    
    # if folder name has a value
    # and the folder name not blank
    # and the folder does NOT yet exist
    if (folderNameExpanded != None
    and folderNameExpanded != ""
    and os.path.exists(folderNameExpanded) == False):
        # create the specified folder
        try:
            if containsFileName == True:
                if os.path.dirname(folderNameExpanded) != '':
                    if os.path.exists(os.path.dirname(folderNameExpanded)) == False:
                        os.makedirs(os.path.dirname(folderNameExpanded))
            else:
                os.makedirs(folderNameExpanded)
        except Exception as e:
            errStr = str(e)
            pgmLogger.error(errStr)

    return errStr


# =============================================================================    
# Expand the specified folder path as needed
# =============================================================================    

def getPathExpanded(path,
                    parentPath = '',
                    pgmLogger=None):
    # default the return value
    pathExpanded = path
    # if it even has a value
    if pathExpanded != None and pathExpanded != '':
        # if the home folder is specified
        if pathExpanded.startswith("~") == True:
            # expand the file path with the home folder
            pathExpanded = os.path.expanduser(pathExpanded)
        # split the folder into its drive and tail
        drive, tail = os.path.splitdrive(pathExpanded)
        # if it's a sub-folder
        if drive == '' and tail.startswith("/") == False:
            if parentPath == None:
                parentPath = ''
            pathExpanded = os.path.join(parentPath, pathExpanded)
        # obtain the folder's absolute path
        pathExpanded = os.path.abspath(pathExpanded)
    # return expanded folder path
    return pathExpanded

    
# =============================================================================    
# Unzip a ZIP archive to a target folder
# =============================================================================    

def zip2dir(zipFileName,
            tgtPathName,
            pgmLogger=None):
    
    errStr = None
    
    pgmLogger.info("")
    pgmLogger.info("=================================")
    pgmLogger.info("ZIP to TGT directory expansion...")
    pgmLogger.info("---------------------------------")
    pgmLogger.info("zipFileName: '%s'" % zipFileName)
    pgmLogger.info("tgtPathName: '%s'" % tgtPathName)
    
    zipFileNameExpanded = getPathExpanded(zipFileName)
    tgtPathNameExpanded = getPathExpanded(tgtPathName)
        
    # verify that ZIP file exists
    if not os.path.exists(zipFileNameExpanded):
        errStr = "ZIP file '%s' does NOT exist!" % zipFileNameExpanded
        pgmLogger.error(errStr)
    else:
        # create the target folder(s) if not extant
        createFolderIfNotExist(tgtPathNameExpanded, containsFileName=False) 
    
        # extract all of the files/folders
        # within the zip file to the target folder
        with zipfile.ZipFile(zipFileNameExpanded, "r") as z:
            z.extractall(tgtPathNameExpanded)
        print ('tgtPathName: "%s"' % tgtPathNameExpanded)

    return tgtPathNameExpanded, errStr

    
# =============================================================================    
# Unzip a ZIP archive to a target folder
# =============================================================================    

def py7zip2dir(archive7z,
               tgtPathName,
               pgmLogger=None,
               isTestMode=False):
    
    errStr = None
    
    pgmLogger.info("")
    pgmLogger.info("===================================")
    pgmLogger.info("7-ZIP to TGT directory expansion...")
    pgmLogger.info("===================================")
    pgmLogger.info("tgtPathName: '%s'" % tgtPathName)
    
    tgtPathNameExpanded = getPathExpanded(tgtPathName)

    # create the target folder(s) if not extant
    createFolderIfNotExist(tgtPathNameExpanded, containsFileName=False) 
    
    # obtain the archive's manifest of file names
    fileNameManifestList = archive7z.getnames() 
    for fileName in fileNameManifestList:
        outfilename = os.path.join(tgtPathNameExpanded, fileName)
        outdir = os.path.dirname(outfilename)
        if not os.path.exists(outdir):
            os.makedirs(outdir)
        with io.open(outfilename, 'wb') as outfile:
            outfile.write(archive7z.getmember(fileName).read())
        print ('outFileName: "%s"' % outfilename)
                    
    return tgtPathNameExpanded, errStr


# =========================
# execute the "main" method
# =========================

if __name__ == "__main__":
    main()
