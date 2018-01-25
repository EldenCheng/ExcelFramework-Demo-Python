from PIL import Image, ImageEnhance
import pytesseract

im = Image.open(r".\Webelementsnapshot.jpg")
imgry = im.convert('L')  # 图像加强，二值化
sharpness = ImageEnhance.Contrast(imgry)  # 对比度增强
sharp_img = sharpness.enhance(2.0)
#sharp_img.show()
#sharp_img.save(r".\Webelementsnapshot2.jpg")


code = pytesseract.image_to_string(im)  # code即为识别出的图片数字str类型
print("The code is %s" % code)
