#!/bin/bash
FILE="/home/rpa01/rpa_naver_brandmall/naverdaily.py"
killall python3;
pkill chrome;
pkill chromedriver;
python3 $FILE $1;
sleep 10s;
