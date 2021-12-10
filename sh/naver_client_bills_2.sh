#!/bin/bash
SHOPS=("ddp" "daeguOut" "jdong" "songdo" "cchung" "daegu" "gasan" "busan" "sinchon" "gimpo" "head")
FILE="/home/rpa01/rpa_naver_brandmall/naver_client_bill.py"
for SHOP in ${SHOPS[@]};
do
    killall python3;
    pkill chrome;
    pkill chromedriver;
    python3 $FILE $SHOP;
    sleep 10s;
done