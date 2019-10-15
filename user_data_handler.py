import cryptography, os, subprocess
from cryptography.fernet import Fernet
from lxml import etree as ET
from io import BytesIO

key = b'b0o1r-lJ_U8vrEnftjOnuJB1D-qPUE9aKtQ33XtFt-E='
fernet = Fernet(key)

def encrypt_user_data(udef_xml):
        root = ET.Element("main")
        email = ET.SubElement(root, "email")

        smtp = ET.SubElement(email, "smtp", port="")
        addr = ET.SubElement(email, "addr")
        passwd = ET.SubElement(email, "passwd")
        recipient = ET.SubElement(email, "recipient")
        _from = ET.SubElement(email, "from")

        services = ET.SubElement(root, "services")
        service = ET.SubElement(services, "service")

        open_new = ET.ElementTree(root)
        with open(udef_xml, "wb") as file:
            file.write(fernet.encrypt(ET.tostring(open_new, pretty_print=True)))
        # subprocess.check_call(["attribs", "+H", udef_xml])

def decrypt_user_data(file):
    global email_user, email_smtp, email_port, email_from, email_passwd, email_recipient, services_list
    file = file

    if not os.path.isfile(file):
        encrypt_user_data(file)

    with open(file, "rb") as f:
        data = fernet.decrypt(f.read())

    tree = ET.parse(BytesIO(data))
    root = tree.getroot()

    email_user = root[0][1].text
    email_smtp = root[0][0].text
    email_port = str(root[0][0].get("port"))
    email_from = root[0][3].text
    email_passwd = root[0][2].text
    email_recipient = []
    services_list = []

    for x in root.findall('email/recipient'):
        email_recipient.append(x.text)

    for x in root.findall('services/service'):
        services_list.append(x.text)
