import user_data_handler as decrypt
import os

cwd = os.getcwd()

udef = "ServiceChecker.conf"

init_decrypt = decrypt.decrypt_user_data(udef)
email_smtp = decrypt.email_smtp
email_port = decrypt.email_port
email_from = decrypt.email_from
email_user = decrypt.email_user
email_passwd = decrypt.email_passwd
email_recipient = decrypt.email_recipient
services = decrypt.services_list

# log_name = f"logfile{today}.txt"
# log_file =
