from string import Template

document_template_html = "template_2.html"

def getFileContent(file_path):
    """
    with open(file_path, "rb") as file_content:
        return file_content
    """
    f = open(file_path, "r")
    #print(f.read())
    return f.read()

html_template_code = getFileContent(document_template_html)

html_template_code = html_template_code.replace("{","$")
html_template_code = html_template_code.replace("}","")
#s = Template('$who likes $what')
#s = s.substitute(who='tim', what='kung pao')
s = Template(html_template_code)
# s = s.substitute(name='tim', greetings='kung pao')

context = { 'name' : "luca" ,"greetings":"nice dick bro","unit":"345sdfv","arrival":"14/09/2020","deperture":"15/09/2023","saludation":"hasta la proximaa","contact":"Pam Martin, REALTOR","company_name":"Pam Martin - Keller Williams","phone":"(251)269-8864 or (251)279-0717","qrcode":"mycode"}
"""
s = s.substitute(name = info_name ,
                greetings = info_greeting,
                unit = info_unit,
                arrival = info_arrival,
                deperture = info_deperture,
                saludation = info_saludation,
                contact = info_contact,
                company_name = info_company_name,
                phone = info_phone,
                qrcode = info_qrcode)
                """
s = s.substitute(context)
print(s)