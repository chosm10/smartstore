#!/bin/bash
SHOPS=("ddp" "decube" "chunho" "daeguOut" "kintex" "pangyo" "donggu" "head" "sinchon" "mdong" "ulsan" "garden" "jdong" "songdo" "daegu" "gasan")
FILE="/home/rpa01/rpa_naver_brandmall/navercarryover.py"
for SHOP in ${SHOPS[@]};
do
    killall python3;
    pkill chrome;
    pkill chromedriver;
    python3 $FILE $SHOP;
    sleep 10s;
done