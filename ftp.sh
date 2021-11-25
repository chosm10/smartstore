#!/bin/bash
address=10.103.200.52
id=rpa
passwd=1234
ftp -n -v $address<<EOF
user $id $passwd
passive
put $1
