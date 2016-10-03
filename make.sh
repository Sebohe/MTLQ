#!/bin/bash
rm -r __pycache__
rm -r build
rm -r dist
pyinstaller --noconsole --onefile --specpath=dist/ MTLQ.py
mv dist/MTLQ.exe .
