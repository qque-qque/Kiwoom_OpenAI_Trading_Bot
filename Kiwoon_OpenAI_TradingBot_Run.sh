#!/bin/bash

# 이미 설치된 라이브러리는 다시 설치하지 않도록 처리
pip install --no-deps -r requirements.txt

# Python 스크립트 실행
py ./src/main.py
