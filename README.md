# extract-form
## Introduction
        In this project, I use Python to extract the information in the form from an image file.
        My method is divided into following steps:<br>
        1. Open an image file<br>
        2. Convert the image to grayscale<br>
        3. Convert the image in grayscale to binaryzation<br>
        4. Erode the binary image to find the perpendicular lines and the transeverse lines in it and then find their intersections which are the vertexes of the cells in the form<br>
        5. Use the intersections to get the information in the form<br>
        6. Write the information of the form in an Excel file<br>
## Python Package
        1. Pillow<br>
        2. numpy<br>
        3. pytesseract<br>
        4. xlwt<br>
## Requirement
        I'm new to Python programming. I hope that you can point out the wrong thing in my code or my method and help me to achieve the function. Thank you!
