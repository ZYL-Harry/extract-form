from PIL import Image
import numpy
import pytesseract
import xlwt

# create an Excel file
count = 0
rownumber = 0
text = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = text.add_sheet('title', cell_overwrite_ok=True)
savepath = 'save_path/excel.xls'

# open the image file
image = numpy.array(Image.open("path/photo.png"))

# convert the picture to grayscale by the average value method
image_array = numpy.array(image)
image_gray = image_array.sum(axis=2)/3

# convert the picture in grayscale to binaryzation with the threshold value "127"
image_gray_array = numpy.array(image_gray)
image_two = numpy.where(image_gray_array < 127, 0, 255)

# erode the picture in binaryzation to find the vertexes of the cells in the table
image_two_array = numpy.array(image_two)
H, W = image_two.shape                      # the height and the width of the picture after binaryzation
image_two_array1 = numpy.array([[numpy.zeros((49, W))], [image_two_array], [numpy.zeros((49, W))]])
image_two_array2 = numpy.array([numpy.zeros(((H+98), 49)), image_two_array1, numpy.zeros(((H+98), 49))])    # change the binary image matrix to prepare for eroding the binary picture
# perpendicular line
compare_array1 = numpy.array([numpy.zeros((99, 49)), numpy.ones((99, 1)), numpy.zeros((99, 49))])           # define structural elements for image corrosion to find perpendicular lines
image_two_row = numpy.array(image_two_array2)
for i in range(50, H+49):                   # get the perpendicular lines by eroding the binary image with the center of each pixels of it
    for j in range(50, W+49):
        image_corrosion_row_point = (image_two_row[i-50:i+48, j-50:j+48]*compare_array1).sum
        if image_corrosion_row_point < 255*99:
            image_two_row = 0
image_corrosion_row = image_two_row[49:(H+48), 49:(W+48)]   # get the matrix only left with perpendicular lines after eroding
# transverse line
compare_array2 = numpy.array([[numpy.zeros((49, 99))], [numpy.ones((1, 99))], [numpy.zeros((49, 99))]])     # define structual elements for image corrosion to find transverse lines
image_two_col = numpy.array(image_two_array2)
for i in range(50, H+49):                   # get the transeverse lines by eroding the binary image with the center of each pixels
    for j in range(50, W+49):
        image_corrosion_point = (image_two_col[i-50:i+48, j-50:j+48]*compare_array2).sum
        if image_corrosion_point < 255*99:
            image_two_col = 0
image_corrosion_col = image_two_col[49:(H+48), 49:(W+48)]   # get the matrix only left with transeverse lines after eroding
# find the intersection(the vertexes of the cells in the table) and then put them in order
intersecting_coordinate = []
for x in range(1, H):           # get the coordinates of intersections in rows and columns and then create a matrix
    list_coordinate = []
    for y in range(1, W):
        if image_corrosion_row[x-1][y-1] == image_corrosion_col[x-1][y-1] == 255:   # judge the pixels where perpendicular lines and transeverse lines coincide to find the intersections
            list_coordinate = numpy.append(list_coordinate, (x, y))
    intersecting_coordinate = numpy.append(intersecting_coordinate, list_coordinate, axis=0)

# create a set which includes the vertexes of cells in the table
rectunitrow = []
rectunit = []
HI, WI = intersecting_coordinate.shape  # get the height and width of the matrix of the intersections
for n in range(1, HI, 2):
    for m in range(1, WI, 2):
        rectunitrow = numpy.append(rectunitrow, (intersecting_coordinate[n-1][m-1], intersecting_coordinate[n-1][m], \
                                                 intersecting_coordinate[n][m-1], intersecting_coordinate[n][m]))    # add coordinates of the four vettexes of the each cell in the table to the matrix in each row
    rectunit = numpy.append(rectunitrow, rectunitrow, axis=0)

# get the information in the cells in the table
RH, RW = rectunit.shape
allline = []
for u in range(1, RH):
    perline = []
    for p in range(1, RW):
        rectunit1 = rectunit[i-1]
        data = image_gray[rectunit1[1]:rectunit1[3], rectunit1[2]:rectunit1[4]]       # get each cells of the picture in grayscale through the vertexes of each cells
        text = pytesseract.image_to_string(data, lang='chi_sim')                      # read the information of each cells
        perline = numpy.append(perline, text)
    if "keyword" in perline:            # judge whether the information in a row should be extracted through the keyword "keywords"
        allline = numpy.append(allline, perline, axis=0)
AH,AW = allline.shape
for AH in range(1, AH):
    count = count + 1
    for AW in range(1, AW):
        rownumber = rownumber+1
        text=allline[AH-1][AW-1]
        sheet.write(count, rownumber, text)         # write the information in the Excel file created in rows and columns

# save the Excel file
text.save(savepath)