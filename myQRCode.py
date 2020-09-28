def qrcode_create(something):
    # pip install pillow
    # pip install qrcode
    # https://towardsdatascience.com/generate-qrcode-with-python-in-5-lines-42eda283f325
    import qrcode
    # Link for website
    input_data = "https://towardsdatascience.com/face-detection-in-10-lines-for-beginners-1787aa1d9127"
    input_data = str(something)
    #Creating an instance of qrcode
    qr = qrcode.QRCode(
            version=1,
            box_size=10,
            border=5)
    qr.add_data(input_data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    image_file = 'qrcode001.png'
    img.save(image_file)
    return image_file