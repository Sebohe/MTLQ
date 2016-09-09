#!/bin/bash

pyinstaller --onedir  --onefile --specpath dist MTLQ.py
mv dist/MTLQ.exe .
