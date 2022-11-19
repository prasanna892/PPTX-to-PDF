import comtypes.client, os

class PPTtoPDF():
    def __init__(self):
        # Creating powerpoint object
        self.powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        
        # Getting current position
        self.current_pos = (self.powerpoint.left, self.powerpoint.top)

        # Hiding powerpoint window from screen
        self.powerpoint.left = -self.powerpoint.Width
        self.powerpoint.top = -self.powerpoint.Height

    def convert(self, inputFileName, outputFileName, deleteFile = False):
        if file.lower().endswith(".pptx"): # Checking if file is .pptx file
            print("Converting : ", file[:-5])

            # Checking if outputFileName has .pdf extension
            outputFileName = outputFileName if outputFileName.lower().endswith(".pdf") else outputFileName + '.pdf'

            # Converting to PDF file
            deck = self.powerpoint.Presentations.Open(inputFileName)
            deck.SaveAs(outputFileName, 32) # 32 for ppt to pdf
            deck.Close()

            print("Conversion completed\n")

            if deleteFile:
                # Permanently delete .pptx file
                os.remove(os.path.join(folder_path, file))

    # Method to close powerpoint
    def close(self):
        # Resetting powerpoint position
        self.powerpoint.left, self.powerpoint.top = self.current_pos

        # Closing powerpoint
        self.powerpoint.Quit()


if __name__ == '__main__':
    pptTOpdf = PPTtoPDF()
    
    folder_path = r"Folder Path"

    # Iterating all files in given folder
    for file in os.listdir(folder_path):
        # Calling PPTtoPDF function
        pptTOpdf.convert(os.path.join(folder_path, file), os.path.join(folder_path, file[:-5]), True)
        
    print("All done...")

    pptTOpdf.close()
