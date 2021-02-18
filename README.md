# extract-form
# Introduction
  In this project, I use Python to extract the information in the form from an image file. My method is divided into following steps:  
    1. Open an image file  
    2. Convert the image to grayscale  
    3. Convert the image in grayscale to binaryzation  
    4. Erode the binary image to find the perpendicular lines and the transeverse lines in it and then find their intersections which are the vertexes of the cells in the form  
    5. Use the intersections to get the information in the form  
    6. Write the information of the form in an Excel file  
