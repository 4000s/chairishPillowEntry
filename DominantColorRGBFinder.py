from PIL import Image,ImageDraw
import webcolors
import RGBtoColorName


def closest_colour(requested_colour):
    min_colours = {}
    for key, name in webcolors.css3_hex_to_names.items():
        r_c, g_c, b_c = webcolors.hex_to_rgb(key)
        rd = (r_c - requested_colour[0]) ** 2
        gd = (g_c - requested_colour[1]) ** 2
        bd = (b_c - requested_colour[2]) ** 2
        min_colours[(rd + gd + bd)] = name
    return min_colours[min(min_colours.keys())]

def get_colour_name(requested_colour):
    try:
        closest_name = actual_name = webcolors.rgb_to_name(requested_colour)
    except ValueError:
        closest_name = closest_colour(requested_colour)
        actual_name = None
    return actual_name, closest_name
def get_colors(infile, outfile, numcolors=5, swatchsize=20, resize=150):

    image = Image.open(infile)
    image = image.resize((resize, resize))
    result = image.convert('P', palette=Image.ADAPTIVE, colors=numcolors)
    result.putalpha(0)
    colors = result.getcolors(resize*resize)

    # Save colors to file

    pal = Image.new('RGB', (swatchsize*numcolors, swatchsize))
    draw = ImageDraw.Draw(pal)

    colorNameList = []
    posx = 0
    for count, col in colors:
        draw.rectangle([posx, 0, posx+swatchsize, swatchsize], fill=col)
        posx = posx + swatchsize
        print(col)
        colorNameList.append(RGBtoColorName.findNearestColorName(col[0],col[1],col[2],RGBtoColorName.WebColorMap))

    print(colorNameList)
    del draw
    pal.save(outfile, "PNG")
    return colorNameList

def dominantColorsSet(image):
    im = Image.open(image)
    w, h = im.size
    print(w,h)
    #im = im.crop((200, 100, 1000, 850))    yourImage.crop((0, 30, w, h-30)).save(...)
    im=im.crop(((w*0.25),(h*0.15),w*0.75,h*0.85))
    im.save('_0.png')
    #deneme

    colorNameList=[]
    colorNameList=get_colors('_0.png', 'outfile.png')
    distinctColorNameSet=set(colorNameList)
    print(distinctColorNameSet)
    return distinctColorNameSet

#dominantColorsSet('C:\\Users\\hasan\\Desktop\\deneme_woveny\\50017.jpg')
