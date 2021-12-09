#!/bin/bash
SHOPS=("mdong" "ulsan" "garden" "jdong" "songdo" "cchung" "daegu" "gimpo" "gasan")
FILE="/home/rpa01/rpa_naver_brandmall/naversettle.py"
for SHOP in ${SHOPS[@]};
do
    killall python3;
    pkill chrome;
    pkill chromedriver;
    python3 $FILE $SHOP;
    sleep 10s;
done