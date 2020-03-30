# -*- coding: utf-8 -*-
"""
Creating the docx file in memory
"""

from io import BytesIO
from zipfile import ZipFile
from shutil import copyfileobj

class DocxFile():
    """
    Creates a docx file in memory.
    """
    
    def __init__(self):
        """
        Initialises a file in memory and turns it into a zipped file
        """
        self.file_in_memory = BytesIO()
        self.zipped_file = ZipFile(self.file_in_memory, 'w')
        
    def write(self, target, string):
        """
        Writes string to given target within zipfile
        """
        if type(string) is not bytes:
            string = string.encode("utf-8")
            
        self.zipped_file.writestr(target, string)
        
    def save(self, path):
        """
        Saves the zip file in path
        """
        self.zipped_file.close()
        self.file_in_memory.seek(0)
        
        
        with open(path, 'wb') as target_file:
            copyfileobj(self.file_in_memory, target_file)
    
    def download(self):
        """
        Returns the file for download
        """
        self.zipped_file.close()
        self.file_in_memory.seek(0)
        
        return self.file_in_memory
    
    def close(self):
        """
        Closes the zip file and deletes the file in memory
        """
        self.zipped_file.close()
        self.file_in_memory.truncate(0)
        self.file_in_memory.seek(0)
        

    
        
    