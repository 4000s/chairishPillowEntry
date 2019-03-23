"""
Original Author  Ernesto P. Adorio, Ph.D
Original Source: http://my-other-life-as-programmer.blogspot.com/2012/02/python-finding-nearest-matching-color.html
Modifed By: JDiscar
This class maps an RGB value to the nearest color name it can find. Code is modified to include
ImageMagick names and WebColor names.
1. Modify the minimization criterion to use least sum of squares of the differences.
2. Provide error checking for input R, G, B values to be within the interval [0, 255].
3. Provide different ways to specify the input RGB values, aside from the (R, G, B) values as done in the program above.
"""

"""
black
blue
brown
copper
gold
gray
green
orange
pink
purple
red
silver
tan
turquiose
white
yellow
"""





WebColorMap = {}
WebColorMap["Sky Blue"] = "#F0F8FF" #Sky Blue
WebColorMap["Light Pink"] = "#FAEBD7" #Light Pink
WebColorMap["Aqua"] = "#00FFFF"
WebColorMap["Aqua"] = "#7FFFD4" # Aqua
WebColorMap["Sky Blue"] = "#F0FFFF" #Sky Blue
WebColorMap["Beige"] = "#F5F5DC"
WebColorMap["Light Pink"] = "#FFE4C4" #Light Pink
WebColorMap["Black"] = "#000000"
WebColorMap["Light Pink"] = "#FFEBCD" #Light Pink
WebColorMap["Blue"] = "#0000FF"
WebColorMap["Violet"] = "#8A2BE2" #Violet
WebColorMap["Brown"] = "#A52A2A"
WebColorMap["Cinnamon"] = "#DEB887" #Cinnamon
WebColorMap["Royal Blue"] = "#5F9EA0" #Royal Blue
WebColorMap["Yellow"] = "#7FFF00" #Yellow
WebColorMap["Brown"] = "#D2691E" #Brown
WebColorMap["Orange"] = "#FF7F50" #Orange
WebColorMap["Royal Blue"] = "#6495ED"#Royal Blue
WebColorMap["Vanilla"] = "#FFF8DC" #Vanilla
WebColorMap["Red"] = "#DC143C" #Red
WebColorMap["Aqua"] = "#00FFFF" #Aqua
WebColorMap["Navy"] = "#00008B" #Navy Blue
WebColorMap["Teal"] = "#008B8B" #Teal
WebColorMap["Chestnut"] = "#B8860B" #Chestnut
WebColorMap["Dark Gray"] = "#A9A9A9" #Dark Gray
WebColorMap["Dark Gray"] = "#A9A9A9" #Dark Gray
WebColorMap["Dark Green"] = "#006400" #Dark Green
WebColorMap["Khaki"] = "#BDB76B" #Khaki
WebColorMap["Magenta"] = "#8B008B" #Magenta
WebColorMap["Dark Green"] = "#556B2F"  #Dark Green
WebColorMap["Orange"] = "#FF8C00" #Orange
WebColorMap["Magenta"] = "#9932CC" #Magenta
WebColorMap["Ruby Red"] = "#8B0000" #Ruby Red
WebColorMap["Salmon"] = "#E9967A" #Salmon
WebColorMap["Green"] = "#8FBC8F" #Green
WebColorMap["Navy"] = "#483D8B" #Navy
WebColorMap["Dark Gray"] = "#2F4F4F" #Dark Gray
WebColorMap["Dark Gray"] = "#2F4F4F" #Dark Gray
WebColorMap["Turquoise"] = "#00CED1"  #Turquoise
WebColorMap["Violet"] = "#9400D3" #Violet
WebColorMap["Bright Pink"] = "#FF1493" #Bright Pink
WebColorMap["Sky Blue"] = "#00BFFF" #Sky Blue
WebColorMap["Gray"] = "#696969" #Gray
WebColorMap["Gray"] = "#696969" #Gray
WebColorMap["Sky Blue"] = "#1E90FF" #Sky Blue
WebColorMap["Brick Red"] = "#B22222" #Brick Red
WebColorMap["White"] = "#FFFAF0" #White
WebColorMap["Green"] = "#228B22" #Green
WebColorMap["Pink"] = "#FF00FF" #Pink
WebColorMap["Gray"] = "#DCDCDC" #Gray
WebColorMap["White"] = "#F8F8FF" #White
WebColorMap["Gold"] = "#FFD700" #Gold
WebColorMap["Gold"] = "#DAA520" #Gold
WebColorMap["Gray"] = "#808080" #Gray
WebColorMap["Gray"] = "#808080" #Gray
WebColorMap["Green"] = "#008000"
WebColorMap["Bright Green"] = "#ADFF2F" #Bright Green
WebColorMap["White"] = "#F0FFF0" #White
WebColorMap["Bright Pink"] = "#FF69B4" #Bright Pink
WebColorMap["Ruby Red"] = "#CD5C5C" #Ruby Red
WebColorMap["Royal Blue"] = "#4B0082" #Royal Blue
WebColorMap["Ivory"] = "#FFFFF0"
WebColorMap["Khaki"] = "#F0E68C"
WebColorMap["Lavender"] = "#E6E6FA"
WebColorMap["Lavender"] = "#FFF0F5" #Lavender
WebColorMap["Bright Green"] = "#7CFC00" #Bright Green
WebColorMap["Beige"] = "#FFFACD" #Beige
WebColorMap["Sky Blue"] = "#ADD8E6" #Sky Blue
WebColorMap["Raspbery Pink"] = "#F08080" #Raspbery Pink
WebColorMap["Turquoise"] = "#E0FFFF" #Turquoise
WebColorMap["Beige"] = "#FAFAD2"  #Beige
WebColorMap["Gray"] = "#D3D3D3" #Gray
WebColorMap["Gray"] = "#D3D3D3" #Gray
WebColorMap["Green"] = "#90EE90" #Green
WebColorMap["Pink"] = "#FFB6C1" #Pink
WebColorMap["Scarlet"] = "#FFA07A" #Scarlet
WebColorMap["Kelly Green"] = "#20B2AA" #Kelly Green
WebColorMap["Sky Blue"] = "#87CEFA" #Sky Blue
WebColorMap["Gray"] = "#778899" #Gray
WebColorMap["Gray"] = "#778899" #Gray
WebColorMap["Cerulean"] = "#B0C4DE" #Cerulean
WebColorMap["Light Yellow"] = "#FFFFE0" #Light Yellow
WebColorMap["Neon Green"] = "#00FF00" #Neon Green
WebColorMap["Kelly Green"] = "#32CD32" #Kelly Green
WebColorMap["Linen"] = "#FAF0E6" #Linen
WebColorMap["Magenta"] = "#FF00FF" #Magenta
WebColorMap["Maroon"] = "#800000" #Maroon
WebColorMap["Cerulean"] = "#66CDAA"  #Cerulean
WebColorMap["Navy Blue"] = "#0000CD" #Navy Blue
WebColorMap["Magenta"] = "#BA55D3" # Magenta
WebColorMap["Purple"] = "#9370D8" #Purple
WebColorMap["Green"] = "#3CB371" #Green
WebColorMap["Blue"] = "#7B68EE" #Blue
WebColorMap["Light Green"] = "#00FA9A" #Light Green
WebColorMap["Turquoise"] = "#48D1CC"  #Turquoise
WebColorMap["Violet"] = "#C71585" #Violet
WebColorMap["Ink Blue"] = "#191970" #Ink Blue
WebColorMap["Cream"] = "#F5FFFA" #Cream
WebColorMap["Rose"] = "#FFE4E1" #Rose
WebColorMap["Mustard"] = "#FFE4B5" #Mustard
WebColorMap["Mustard"] = "#FFDEAD"  #Mustard
WebColorMap["Navy Blue"] = "#000080" #Navy Blue
WebColorMap["Bubble Gum"] = "#FDF5E6" #Bubble Gum
WebColorMap["Green"] = "#808000" #Green
WebColorMap["Green"] = "#6B8E23" #Green
WebColorMap["orange"] = "#FFA500" #orange
WebColorMap["Tangerine"] = "#FF4500" #Tangerine
WebColorMap["Purple"] = "#DA70D6" #Purple
WebColorMap["Khaki"] = "#EEE8AA" #Khaki
WebColorMap["Light Green"] = "#98FB98" #Light Green
WebColorMap["Turquoise"] = "#AFEEEE" #Turquoise
WebColorMap["Violet"] = "#D87093" #Violet
WebColorMap["Bubble Gum"] = "#FFEFD5" #Bubble Gum
WebColorMap["Bubble Gum"] = "#FFDAB9" #Bubble Gum
WebColorMap["Brown"] = "#CD853F" #Brown
WebColorMap["Pink"] = "#FFC0CB" #Pink
WebColorMap["Purple"] = "#DDA0DD" #Purple
WebColorMap["Turquoise"] = "#B0E0E6" #Turquoise
WebColorMap["Purple"] = "#800080"  #Purple
WebColorMap["Red"] = "#FF0000"
WebColorMap["Pink"] = "#BC8F8F" #Ruby Red
WebColorMap["Royal Blue"] = "#4169E1" #Royal Blue
WebColorMap["Brown"] = "#8B4513" #Brown
WebColorMap["Raspberry Red"] = "#FA8072" #Raspberry Red
WebColorMap["Chocolate"] = "#F4A460" #Chocolate
WebColorMap["Forest Green"] = "#2E8B57" #Forest Green
WebColorMap["Bubble Gum"] = "#FFF5EE" #Bubble Gum
WebColorMap["Cinnamon"] = "#A0522D" #Cinnamon
WebColorMap["Silver"] = "#C0C0C0" #Silver
WebColorMap["Sky Blue"] = "#87CEEB" #Sky Blue
WebColorMap["Royal Blue"] = "#6A5ACD" #Royal Blue
WebColorMap["Gray"] = "#708090" #Gray
WebColorMap["Gray"] = "#708090" #Gray
WebColorMap["Ivory"] = "#FFFAFA" #Ivory
WebColorMap["Light Green"] = "#00FF7F" #Light Green
WebColorMap["Royal Blue"] = "#4682B4" #Royal Blue
WebColorMap["Tan"] = "#D2B48C" #Tan
WebColorMap["Kelly Green"] = "#008080" #Kelly Green
WebColorMap["Violet"] = "#D8BFD8" #Violet
WebColorMap["Orange"] = "#FF6347" #Orange
WebColorMap["Turquoise"] = "#40E0D0" #Turquoise
WebColorMap["Violet"] = "#EE82EE" #Violet
WebColorMap["Apricot"] = "#F5DEB3" #Apricot
WebColorMap["White"] = "#FFFFFF" #White
WebColorMap["Gray"] = "#F5F5F5" #Gray
WebColorMap["Yellow"] = "#FFFF00" #Yellow
WebColorMap["Light Green"] = "#9ACD32" #Light Green
WebColorMap["Black"] = "#191919" 
WebColorMap["Gray"] = "#666151"
WebColorMap["Dove Gray"] = "#ab9f84"
WebColorMap["Brown"] = "#694234"

def rgbFromStr(s):
   # s starts with a #.
    r, g, b = int(s[1:3],16), int(s[3:5], 16),int(s[5:7], 16)
    return r, g, b

def findNearestColorName(R,G,B,Map):
    mindiff = None
    for d in Map:
        r, g, b = rgbFromStr(Map[d])
        diff = abs(R -r)*256 + abs(G-g)* 256 + abs(B- b)* 256
        if mindiff is None or diff < mindiff:
            mindiff = diff
            mincolorname = d
    return mincolorname

#print(findNearestColorName(34, 34, 33,WebColorMap))
