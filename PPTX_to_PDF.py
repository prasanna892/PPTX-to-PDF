import comtypes.client, os

def PPTtoPDF(inputFileName, outputFileName):
    # Creating powerpoint object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    
    # Hiding powerpoint window from screen
    powerpoint.left = -powerpoint.Width
    powerpoint.top = -powerpoint.Height

    # Checking if outputFileName has .pdf extension
    outputFileName = outputFileName if '.pdf' in outputFileName else outputFileName + '.pdf'

    # Opening powerpoint and converting to PDF file
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, 32) # 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()

if __name__ == '__main__':
    
    folder_path = r"Folder Path"

    for file in os.listdir(folder_path):
        if file.endswith(".pptx") or file.endswith(".PPTX"):
            print("Converting : ", file[:-5])
            # Calling PPTtoPDF function
            PPTtoPDF(os.path.join(folder_path, file), os.path.join(folder_path, file[:-5]))
            print("Conversion completed\n")
            os.remove(os.path.join(folder_path, file)) # Permanently delete .pptx file
        
    print("All done...")