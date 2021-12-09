#!/bin/bash
SHOPS=("muyeog" "donggu" "pangyo")
FILE="/home/rpa01/rpa_naver_brandmall/naver_client_bill.py"
for SHOP in ${SHOPS[@]};
do
    killall python3;
    pkill chrome;
    pkill chromedriver;
    python3 $FILE $SHOP;
    sleep 10s;
done