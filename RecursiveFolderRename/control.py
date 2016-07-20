import os
#rootDir = '/media/nyeinchan/601C16171C15E8B4/Development/Ionic/hioutlet'
rootDir = '/media/nyeinchan/601C16171C15E8B4/Development/Git/RecursiveFolderRename/telecom'
def getList(path):
    for root, dirs, files in os.walk(path):
        for d in dirs:
            yield os.path.basename(os.path.normpath(os.path.join(root, d)))


def generateList(path):
    for dirName in getList(path):
        print dirName

if __name__ == '__main__':
    generateList(rootDir)
