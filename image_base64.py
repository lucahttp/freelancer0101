import base64
attachement = "qrcode001.png"
# https://stackoverflow.com/questions/3715493/encoding-an-image-file-with-base64
"""
with open(attachement, "rb") as image_file:
    encoded_image = base64.b64encode(image_file.read())
"""
 
def get_base64_encoded_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode('utf-8')
# encoded_image = base64.b64encode(attachement.getvalue()).decode("utf-8")

encoded_image = get_base64_encoded_image(attachement)
print(encoded_image)
qrcode_html = '<img src="data:image/png;base64,%s"/>' % encoded_image

print(qrcode_html)
#print(qrcode_html)