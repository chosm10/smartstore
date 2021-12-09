#!/bin/bash
SHOPS=("ddp" "gimpo" "garden" "gasan" "daeguOut")
FILE="/home/rpa01/rpa_naver_brandmall/naverdaily.py"
for SHOP in ${SHOPS[@]};
do
    killall python3;
    pkill chrome;
    pkill chromedriver;
    python3 $FILE $SHOP;
    sleep 10s;
done