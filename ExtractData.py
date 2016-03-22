#-------------------------------------------------------------------------------
# Name:        Extract Ag Data
# Purpose:     Pull Ag and Base Data, Convert, Zip, Post, Email Group
#
# Author:      grillb
#
# Created:     06/05/2015
#-------------------------------------------------------------------------------
#Modules
import arcpy, os, sys, zipfile, ftplib, time, win32com.client
from datetime import datetime
now = datetime.now()

#$$$$$$$$$$Global Admin Section$$$$$$$$$$
# Set Workspace
root = r"C:\TEMP"
TempOutput = os.path.join(root, "AgData_%s%s%s_%s%s") % (now.month, now.day, now.year, now.hour, now.minute)
TempOutputApps = os.path.join(root, "AgData2_%s%s%s_%s%s") % (now.month, now.day, now.year, now.hour, now.minute) #has apps data in there

#Create Workspace Folders
if not os.path.exists(root): os.makedirs(root)

#Include Ag Applicant data?
rightanswer = 0
while rightanswer < 1:
    includeapps = raw_input("Include Ag Applicants? (Y or N):")
    if includeapps == "Y" or includeapps == "N":
        rightanswer += 2

#$$$$$$$$$$Function Section$$$$$$$$$$
#Extract the Data
def extractSection():
    #Copy From
    SDE_parcel = "\\path\\to\\data\\parcel_poly"
    SDE_ease = "\\path\\to\\data\\AgEasements"
    SDE_asa = "\\path\\to\\data\\Agsecurity"
    SDE_apps = "\\path\\to\\data\\AgApplications"
    SDE_muni = "\\path\\to\\data\\Muni"
    SDE_uga = "\\path\\to\\data\\Uga_poly"
    SDE_zoning = "\\path\\to\\data\\Zoning"
    SDE_water = "\\path\\to\\data\\Water_service_areas"
    SDE_waste = "\\path\\to\\data\\Wastewater_service_areas"
    SDE_roads = "\\path\\to\\data\\RDCLINE"

    agdata_apps = [SDE_parcel, SDE_ease, SDE_asa, SDE_apps, SDE_muni, SDE_uga, SDE_zoning, SDE_water, SDE_waste, SDE_roads]

    agdata = [SDE_parcel, SDE_ease, SDE_asa, SDE_muni, SDE_uga, SDE_zoning, SDE_water, SDE_waste, SDE_roads]

    # Fields to delete
    deleteFields = "FIELDS;TO;DELETE"

    #Convert to shpfile, drop unnecessary or sensetive fields
    def Convert(output, aglist):
        if not os.path.exists(output): os.makedirs(output)
        arcpy.FeatureClassToShapefile_conversion(aglist, output)
        shp_ease = os.path.join(output, "AgEasements.shp")
        shp_asa = os.path.join(output, "Agsecurity.shp")
        shp_apps = os.path.join(output, "AgApplications.shp")
        arcpy.DeleteField_management(shp_asa, deleteFields)
        arcpy.DeleteField_management(shp_ease, deleteFields)

    #Run Convert function, decide whether Apps are included
    if includeapps == "Y":
        Convert(TempOutputApps, agdata_apps)
        arcpy.DeleteField_management(os.path.join(TempOutputApps, "AgApplications.shp"), deleteFields)
        if not os.path.exists(TempOutput): os.makedirs(TempOutput)
        arcpy.Copy_management(TempOutputApps, TempOutput)
        arcpy.Delete_management(os.path.join(TempOutput, "AgApplications.shp"))
    else:
        Convert(TempOutput, agdata)

#Zip the Data
def zipSection():
    def zipdir(dirPath=None, zipFilePath=None, includeDirInZip=True):

        if not zipFilePath:
            zipFilePath = dirPath + ".zip"
        if not os.path.isdir(dirPath):
            raise OSError("dirPath argument must point to a directory. "
                "'%s' does not." % dirPath)
        parentDir, dirToZip = os.path.split(dirPath)

        #funtion to prepare archieve path
        def trimPath(path):
            archivePath = path.replace(parentDir, "", 1)
            if parentDir:
                archivePath = archivePath.replace(os.path.sep, "", 1)
            if not includeDirInZip:
                archivePath = archivePath.replace(dirToZip + os.path.sep, "", 1)
            return os.path.normcase(archivePath)

        outFile = zipfile.ZipFile(zipFilePath, "w", compression=zipfile.ZIP_DEFLATED)
        for (archiveDirPath, dirNames, fileNames) in os.walk(dirPath):
            for fileName in fileNames:
                filePath = os.path.join(archiveDirPath, fileName)
                outFile.write(filePath, trimPath(filePath))
            #Make sure we get empty directories as well
            if not fileNames and not dirNames:
                zipInfo = zipfile.ZipInfo(trimPath(archiveDirPath) + "/")
                outFile.writestr(zipInfo, "")
        outFile.close()

    if includeapps == "Y":
        zipdir(TempOutput)
        zipdir(TempOutputApps)
    else:
        print "Zipping:", TempOutput
        zipdir(TempOutput)


#FTP the Data
def ftpSection():
    #FTP site
    ftpUrl = "ftp address"
    ftp = ftplib.FTP(ftpUrl)

    #Establish Connection
    try:
        ftp.login("user", "pass")
        print "Connected to", ftpUrl
    except Exception:
        print "Invalid credentials"
        sys.exit()

    #Upload selected file, parameters are file and directory, double quotes will put file into root
    def uploadtoFTP(fileName, dirLoc):
        if dirLoc != "":
            ftp.cwd(dirLoc)
        else:
            ftp.cwd("/")

        fileBase = os.path.basename(fileName)

        data = []

        ftp.dir(data.append)

        for line in data:
            if fileBase == line.split(" ")[-1]:
                print "The file {} already exists on FTP server".format(line.split(" ")[-1])
                sys.exit()

        def upload():
            print "Uploading:", (fileName)
            ext = os.path.splitext(fileName)[1]
            if ext in (".txt", ".htm", ".html"):
                ftp.storlines("STOR " + os.path.basename(fileName), open(fileName,"r"))
            else:
                ftp.storbinary("STOR " + os.path.basename(fileName), open(fileName, "rb"), 1024)

        try:
            upload()
            print "{} Uploaded\n".format(fileBase)
        except Exception:
            print "Sorry, something didnt work"
            sys.exit()


    if includeapps == "Y":
        uploadtoFTP(TempOutput + ".zip", "gisdata")
        uploadtoFTP(TempOutputApps + ".zip", "")
    else:
        uploadtoFTP(TempOutput + ".zip", "")

    ftp.quit()



#Email the Users
def mailSection():
    def sendemail(subject, theText, address):
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = subject
        newMail.Body = theText
        newMail.To = ""
        newMail.BCC = address
        newMail.Send()
    #Send it
    orgs = "hi@example.com; you@example.com; "
    appraisers = "hi@example.com; you@example.com "
    surveyors = "hi@example.com; you@example.com "

    if includeapps == "Y":
        #appraisers
        subject = "Data"
        textbody = " Hi, the latest GIS data is now available on our FTP. file: %s.zip \r\n Thanks." % (TempOutputApps[8:])
        sendto = appraisers + orgs
        sendemail(subject, textbody, sendto)
        #surveyors
        subject2 = "Data"
        textbody2 = " Hi, the latest GIS data is now available on our FTP. file: %s.zip \r\n Thanks." % (TempOutput[8:])
        sendto2 = surveyors
        sendemail(subject2, textbody2, sendto2)
    else:
        subject = "Data"
        textbody = " Hi, the latest GIS data is now available on our FTP. file: %s.zip \r\n Thanks." % (TempOutput[8:])
        sendto = appraisers + surveyors + orgs
        sendemail(subject, textbody, sendto)

#$$$$$$$$$$Bring It Together$$$$$$$$$$
def main():
    extractSection()
    zipSection()
    ftpSection()
    mailSection()


#$$$$$$$$$$And Execute It$$$$$$$$$$
main()
