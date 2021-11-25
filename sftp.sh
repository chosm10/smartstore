#!/bin/bash
address=10.103.200.117;
id=rpa;
passwd=1234;
#put 할 때 백그라운드로 전송해야 세그멘테이션 폴트 안남
lftp sftp://$id:$passwd@$address<<EOF
put $1&
quit
EOF