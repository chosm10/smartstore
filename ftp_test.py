# import ftplib
# import os
# session = ftplib.FTP('localhost','chosm','1234')
# session.cwd('./naver')
# file = open('C:/Users/Administrator/Downloads/brands_muyeog.csv','rb')                  # file to send
# session.storbinary('STOR muyeog.csv', file)     # send the file
# file.close()                                    # close file and FTP
# session.quit()

import pysftp
import os

filename = 'file_to_upload.html'
filepath = os.path.join("path", "to", "projectDir")
localFile = os.path.join(filepath, filename)
cnopts = pysftp.CnOpts(knownhosts=os.path.join(filepath, "keyfile"))
host = 'localhost'
user = 'chosm'
pass_ = '1234'

with pysftp.Connection(host=host, username=user, password=pass_, cnopts=cnopts) as sftp:
  sftp.put(localFile, filename)
#   sftp.get(fileOnServer, locationOnPC)