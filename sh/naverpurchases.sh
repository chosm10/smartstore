#!/bin/bash
SHOPS=("ddp" "garden" "daeguOut" "songdo" "mia" "gimpo" "decube" "gasan" "chunho")
FILE="/home/rpa01/rpa_naver_brandmall/naverpurchase.py"
for SHOP in ${SHOPS[@]};
do
    killall python3;
    pkill chrome;
    pkill chromedriver;
    python3 $FILE $SHOP;
    sleep 10s;
done