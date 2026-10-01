#!/bin/sh
python3 -m pip install -r requirements.txt
python3 -m uvicorn app.main:app --reload
