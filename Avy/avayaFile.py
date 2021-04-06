import datetime
import json
import os
import re
import shutil


class Fileop():
    def isDirectory(self, fDir):
        return os.path.isdir(fDir)

    def countDir(self, dPath):
        dirListing = next(os.walk(dPath))[2]
        return len(dirListing)

    def CjsonLoad(self, jfile):
        fdir = os.path.join(Fileop.dwnDir(''), "conf")
        condir = Fileop.isDirectory('', fdir)
        if condir:
            confile = os.path.join(fdir, jfile)
            with open(confile, "r") as j:
                return json.load(j)

    def SjsonLoad(self, jfile):
        fdir = Fileop.dwnDir('')
        condir = Fileop.isDirectory('', fdir)
        if condir:
            confile = os.path.join(fdir, jfile)
            with open(confile, "r") as j:
                return json.load(j)

    def curWorkDir(self):
        return os.path.dirname(os.path.realpath(__file__))

    def makDir(self, Folder):
        try:
            os.makedirs(Folder)
        except OSError as e:
            print("Warning making {0} : MSG - {1}".format(Folder, e))

    def dwnDir(self):
        return os.path.normpath(os.getcwd() + os.sep + os.pardir)

    def newDirec(self, nDirName):
        return os.mkdir(nDirName)

    def RecFileDir(self, dirName):
        new_Folder = os.path.join(Fileop.dwnDir(''), dirName)
        dFlag = Fileop.isDirectory('', new_Folder)
        if not dFlag:
            # make directory
            try:
                Fileop.newDirec('', new_Folder)
            except OSError:
                print("Creation of the directory %s failed. \n" % new_Folder)
            else:
                print("Successfully created the directory %s.\n " % new_Folder)
        else:
            print("Directory ( %s ) already exists.\n" % new_Folder)
        return new_Folder

    def newest(self, path):
        files = os.listdir(path)
        lenfile = len(files)
        if lenfile != 0:
            paths = [os.path.join(path, basename) for basename in files]
            return max(paths, key=os.path.getctime)
        else:
            print("Directory ( %s ) is empty\n", path)

    def removefiles(self, folder):
        files_in_directory = os.listdir(folder)
        filtered_files = [file for file in files_in_directory if file.endswith(".wav")]
        dircount = Fileop.countDir('', folder)
        if dircount > 1:
            for file in filtered_files:
                path_to_file = os.path.join(folder, file)
                os.remove(path_to_file)
        else:
            print('Failed to delete files, {0} is empty: \n'.format(folder))

    def moveFiles(self, froDir, toDir, froFile, toFile):
        try:
            shutil.move(os.path.join(froDir, froFile), os.path.join(toDir, toFile))
            print("Successfully moved {0}.\n".format(os.path.join(froDir, froFile)))
        except OSError:
            print("Could not move file ({0}) operation.\n".format(os.path.join(froDir, froFile)))

    def main(self):
        # Check if directories have been created
        source_param = Fileop.CjsonLoad('', "conf.json")
        source_rep = os.path.join(Fileop.dwnDir(''), "reports")
        Fileop.RecFileDir('', source_rep)
        dwn_dir = source_param['download_dir']
        # Recordings directory based on current date
        recFolder = "Recordings" + datetime.datetime.now().strftime("%Y%m%d")
        dirRecs = Fileop.RecFileDir('', recFolder)
        # print(dirRecs)
        # get latest data report file
        newFile = Fileop.newest('', source_rep)
        # print (newFile)
        count = 0
        if Fileop.countDir('', dwn_dir) != 0:
            with open(newFile, "r") as nf:
                lines = nf.readlines()
                for line in lines:
                    count += 1
                    line_id = ' '.join(re.findall(r'\b\w+\b', line)[:+1])
                    line_data = ' '.join(re.findall(r'\b\w+\b', line)[:]).replace(line_id, "")
                    line_data = "_".join(line_data.split())
                    print("line {0} - file ID : {1} file metadata :- {2} \n".format(count, line_id, line_data))
                    # move and rename files
                    Fileop.moveFiles("", dwn_dir, dirRecs, line_id + ".wav", line_data + ".wav")
        else:
            print("Warning: Call recordings download did not run!\n")
# if __name__ == "__main__":
#     main()
